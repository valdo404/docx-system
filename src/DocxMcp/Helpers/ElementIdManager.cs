using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxMcp.Helpers;

/// <summary>
/// Central service for assigning and reading stable element IDs.
///
/// Two complementary mechanisms:
/// 1. dmcp:id custom attribute on ALL elements — universal coverage, SDK-preserved.
/// 2. w14:paraId native attribute on Paragraphs and TableRows — Word-native, survives Word open/save.
///
/// When Word strips dmcp:id attributes (on save), re-derives them from surviving w14:paraId values.
/// </summary>
public static class ElementIdManager
{
    public const string DmcpNamespace = "http://docx-mcp.dev/id";
    public const string DmcpPrefix = "dmcp";
    private const string W14Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string McNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private static readonly XmlQualifiedName DmcpIdName = new("id", DmcpNamespace);

    /// <summary>
    /// Generate a compliant hex ID: 8-char uppercase hex, value in [1, 0x7FFFFFFF].
    /// </summary>
    public static string GenerateId(HashSet<string>? existing = null)
    {
        string id;
        do
        {
            id = Random.Shared.Next(1, int.MaxValue).ToString("X8");
        } while (existing is not null && !existing.Add(id));

        return id;
    }

    /// <summary>
    /// Ensure the dmcp namespace and mc:Ignorable are declared on the document root
    /// so Word will gracefully skip (and discard) the custom attributes.
    /// </summary>
    public static void EnsureNamespace(WordprocessingDocument doc)
    {
        var document = doc.MainDocumentPart?.Document;
        if (document is null) return;

        // Add namespace declarations using the proper API
        var nsDecls = document.NamespaceDeclarations.ToDictionary(d => d.Key, d => d.Value);

        if (!nsDecls.ContainsKey("mc"))
            document.AddNamespaceDeclaration("mc", McNamespace);

        if (!nsDecls.ContainsKey(DmcpPrefix))
            document.AddNamespaceDeclaration(DmcpPrefix, DmcpNamespace);

        // Ensure mc:Ignorable includes "dmcp"
        // Use GetAttributes() to safely check for the attribute's existence
        var attrs = document.GetAttributes();
        var ignorableAttr = attrs.FirstOrDefault(a => a.LocalName == "Ignorable" && a.NamespaceUri == McNamespace);

        if (ignorableAttr.Value is null || string.IsNullOrEmpty(ignorableAttr.Value))
        {
            document.SetAttribute(new OpenXmlAttribute("mc", "Ignorable", McNamespace, DmcpPrefix));
        }
        else if (!ignorableAttr.Value.Split(' ').Contains(DmcpPrefix))
        {
            document.SetAttribute(new OpenXmlAttribute("mc", "Ignorable", McNamespace,
                ignorableAttr.Value + " " + DmcpPrefix));
        }
    }

    /// <summary>
    /// Assign IDs to ALL elements in the document.
    /// - Paragraphs/TableRows: set both dmcp:id and w14:paraId/textId
    /// - Tables/Cells/Runs/Drawings/Hyperlinks/BookmarkStarts: set dmcp:id only
    /// - Preserves existing IDs; only fills missing ones
    /// - On Word-stripped docs: re-derives dmcp:id from w14:paraId for paragraphs/rows
    /// </summary>
    public static void EnsureAllIds(WordprocessingDocument doc)
    {
        var existing = CollectExistingIds(doc);
        var mainPart = doc.MainDocumentPart;
        if (mainPart?.Document is null) return;

        AssignIdsInPart(mainPart.Document, existing);

        // Scan header/footer parts
        foreach (var hp in mainPart.HeaderParts)
        {
            if (hp.Header is not null)
                AssignIdsInPart(hp.Header, existing);
        }

        foreach (var fp in mainPart.FooterParts)
        {
            if (fp.Footer is not null)
                AssignIdsInPart(fp.Footer, existing);
        }
    }

    /// <summary>
    /// Assign an ID to a single newly-created element.
    /// For Paragraph/TableRow, also sets w14:paraId and w14:textId.
    /// </summary>
    public static void AssignId(OpenXmlElement element, HashSet<string>? existing = null)
    {
        existing ??= [];
        var id = GenerateId(existing);

        SetDmcpId(element, id);

        if (element is Paragraph p)
        {
            if (p.ParagraphId is null || string.IsNullOrEmpty(p.ParagraphId.Value))
                p.ParagraphId = new HexBinaryValue(id);
            if (p.TextId is null || string.IsNullOrEmpty(p.TextId.Value))
                p.TextId = new HexBinaryValue(GenerateId(existing));
        }
        else if (element is TableRow tr)
        {
            if (tr.ParagraphId is null || string.IsNullOrEmpty(tr.ParagraphId.Value))
                tr.ParagraphId = new HexBinaryValue(id);
            if (tr.TextId is null || string.IsNullOrEmpty(tr.TextId.Value))
                tr.TextId = new HexBinaryValue(GenerateId(existing));
        }
    }

    /// <summary>
    /// Read the dmcp:id from any element. Falls back to w14:paraId for Paragraph/TableRow.
    /// </summary>
    public static string? GetId(OpenXmlElement element)
    {
        var dmcpId = GetDmcpId(element);
        if (dmcpId is not null) return dmcpId;

        // Fall back to w14:paraId for paragraphs and rows
        if (element is Paragraph p)
            return p.ParagraphId?.Value;
        if (element is TableRow tr)
            return tr.ParagraphId?.Value;

        return null;
    }

    // --- Internal helpers ---

    private static void AssignIdsInPart(OpenXmlElement root, HashSet<string> existing)
    {
        foreach (var element in root.Descendants())
        {
            if (!IsIdTarget(element)) continue;

            var currentDmcpId = GetDmcpId(element);

            if (element is Paragraph p)
            {
                // If dmcp:id is missing but w14:paraId exists (Word-stripped), re-derive
                if (currentDmcpId is null && p.ParagraphId?.Value is string paraId && paraId.Length > 0)
                {
                    SetDmcpId(element, paraId);
                    existing.Add(paraId);
                }
                else if (currentDmcpId is null)
                {
                    var id = GenerateId(existing);
                    SetDmcpId(element, id);
                    p.ParagraphId = new HexBinaryValue(id);
                    p.TextId = new HexBinaryValue(GenerateId(existing));
                }
                else
                {
                    // dmcp:id exists — ensure w14:paraId matches
                    existing.Add(currentDmcpId);
                    if (p.ParagraphId is null || string.IsNullOrEmpty(p.ParagraphId.Value))
                        p.ParagraphId = new HexBinaryValue(currentDmcpId);
                    if (p.TextId is null || string.IsNullOrEmpty(p.TextId.Value))
                        p.TextId = new HexBinaryValue(GenerateId(existing));
                }
            }
            else if (element is TableRow tr)
            {
                if (currentDmcpId is null && tr.ParagraphId?.Value is string rowParaId && rowParaId.Length > 0)
                {
                    SetDmcpId(element, rowParaId);
                    existing.Add(rowParaId);
                }
                else if (currentDmcpId is null)
                {
                    var id = GenerateId(existing);
                    SetDmcpId(element, id);
                    tr.ParagraphId = new HexBinaryValue(id);
                    tr.TextId = new HexBinaryValue(GenerateId(existing));
                }
                else
                {
                    existing.Add(currentDmcpId);
                    if (tr.ParagraphId is null || string.IsNullOrEmpty(tr.ParagraphId.Value))
                        tr.ParagraphId = new HexBinaryValue(currentDmcpId);
                    if (tr.TextId is null || string.IsNullOrEmpty(tr.TextId.Value))
                        tr.TextId = new HexBinaryValue(GenerateId(existing));
                }
            }
            else
            {
                // All other ID-target elements: just set dmcp:id if missing
                if (currentDmcpId is null)
                {
                    var id = GenerateId(existing);
                    SetDmcpId(element, id);
                }
                else
                {
                    existing.Add(currentDmcpId);
                }
            }
        }
    }

    /// <summary>
    /// Collect all existing IDs (both dmcp:id and w14:paraId) for collision avoidance.
    /// </summary>
    private static HashSet<string> CollectExistingIds(WordprocessingDocument doc)
    {
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var mainPart = doc.MainDocumentPart;
        if (mainPart?.Document is null) return ids;

        CollectIdsFromPart(mainPart.Document, ids);

        foreach (var hp in mainPart.HeaderParts)
        {
            if (hp.Header is not null)
                CollectIdsFromPart(hp.Header, ids);
        }

        foreach (var fp in mainPart.FooterParts)
        {
            if (fp.Footer is not null)
                CollectIdsFromPart(fp.Footer, ids);
        }

        return ids;
    }

    private static void CollectIdsFromPart(OpenXmlElement root, HashSet<string> ids)
    {
        foreach (var element in root.Descendants())
        {
            var dmcpId = GetDmcpId(element);
            if (dmcpId is not null) ids.Add(dmcpId);

            if (element is Paragraph p && p.ParagraphId?.Value is string pid)
                ids.Add(pid);
            if (element is TableRow tr && tr.ParagraphId?.Value is string rid)
                ids.Add(rid);
        }
    }

    private static bool IsIdTarget(OpenXmlElement element) =>
        element is Paragraph or Table or TableRow or TableCell or Run
            or Drawing or Hyperlink or BookmarkStart;

    internal static string? GetDmcpId(OpenXmlElement element)
    {
        // Use GetAttributes() to safely read extended/custom attributes
        // GetAttribute(name, ns) throws KeyNotFoundException for unknown schemas
        var attr = element.GetAttributes()
            .FirstOrDefault(a => a.LocalName == "id" && a.NamespaceUri == DmcpNamespace);
        return string.IsNullOrEmpty(attr.Value) ? null : attr.Value;
    }

    internal static void SetDmcpId(OpenXmlElement element, string id)
    {
        element.SetAttribute(new OpenXmlAttribute(DmcpPrefix, "id", DmcpNamespace, id));
    }
}
