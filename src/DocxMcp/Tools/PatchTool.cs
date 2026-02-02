using System.ComponentModel;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ModelContextProtocol.Server;
using DocxMcp.Helpers;
using DocxMcp.Paths;

namespace DocxMcp.Tools;

[McpServerToolType]
public sealed class PatchTool
{
    [McpServerTool(Name = "apply_patch"), Description(
        "Modify a document using JSON patches (RFC 6902 adapted for OOXML).\n" +
        "Maximum 10 operations per call. Split larger changes into multiple calls.\n\n" +
        "Operations:\n" +
        "  add — Insert element at path. Use /body/children/N for positional insert.\n" +
        "  replace — Replace element or property at path.\n" +
        "  remove — Delete element at path.\n" +
        "  move — Move element from one location to another.\n" +
        "  copy — Duplicate element to another location.\n" +
        "  replace_text — Find/replace text preserving run-level formatting.\n" +
        "  remove_column — Remove a column from a table by index.\n\n" +
        "Paths support stable element IDs (preferred over indices for existing content):\n" +
        "  /body/paragraph[id='1A2B3C4D'] — target paragraph by ID\n" +
        "  /body/table[id='5E6F7A8B']/row[id='AABB1122'] — target row by ID\n\n" +
        "Value types (for add/replace):\n" +
        "  Paragraph with runs (preserves styling):\n" +
        "    {\"type\": \"paragraph\", \"runs\": [{\"text\": \"bold\", \"style\": {\"bold\": true}}, {\"tab\": true}, {\"text\": \"normal\"}]}\n" +
        "    {\"type\": \"paragraph\", \"properties\": {\"alignment\": \"center\", \"tabs\": [{\"position\": 4680, \"alignment\": \"center\"}]}, \"runs\": [...]}\n" +
        "  Paragraph with flat text (legacy):\n" +
        "    {\"type\": \"paragraph\", \"text\": \"...\", \"style\": {\"bold\": true}}\n" +
        "  Heading with runs:\n" +
        "    {\"type\": \"heading\", \"level\": 2, \"runs\": [{\"text\": \"Title \", \"style\": {\"color\": \"2E5496\"}}, {\"tab\": true}, {\"text\": \"Company\"}]}\n" +
        "  Table:\n" +
        "    {\"type\": \"table\", \"headers\": [\"Col1\",\"Col2\"], \"rows\": [[\"A\",\"B\"]]}\n" +
        "    Rich: {\"type\": \"table\", \"headers\": [{\"text\": \"H1\", \"shading\": \"E0E0E0\"}], \"rows\": [{\"cells\": [{\"text\": \"A\", \"style\": {\"bold\": true}}]}]}\n" +
        "  Row: {\"type\": \"row\", \"cells\": [\"A\", \"B\"], \"is_header\": true, \"height\": 400}\n" +
        "  Cell: {\"type\": \"cell\", \"text\": \"val\", \"shading\": \"FF0000\", \"col_span\": 2, \"vertical_align\": \"center\"}\n" +
        "  Other: image, hyperlink, page_break, section_break, list\n\n" +
        "Run content types:\n" +
        "  {\"text\": \"hello\", \"style\": {\"bold\": true, \"italic\": true, \"font_size\": 14, \"font_name\": \"Arial\", \"color\": \"FF0000\"}}\n" +
        "  {\"tab\": true} — tab character\n" +
        "  {\"break\": \"line\"} — line/page/column break\n\n" +
        "Style properties (for replace on /style paths):\n" +
        "  {\"bold\": true, \"italic\": false, \"font_size\": 14, \"color\": \"FF0000\"}\n\n" +
        "replace_text: {\"op\": \"replace_text\", \"path\": \"/body/paragraph[0]\", \"find\": \"old\", \"replace\": \"new\"}\n" +
        "remove_column: {\"op\": \"remove_column\", \"path\": \"/body/table[0]\", \"column\": 2}")]
    public static string ApplyPatch(
        SessionManager sessions,
        [Description("Session ID of the document.")] string doc_id,
        [Description("JSON array of patch operations (max 10 per call).")] string patches)
    {
        var session = sessions.Get(doc_id);
        var wpDoc = session.Document;
        var mainPart = wpDoc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no MainDocumentPart.");

        JsonElement patchArray;
        try
        {
            patchArray = JsonDocument.Parse(patches).RootElement;
        }
        catch (JsonException ex)
        {
            return $"Error: Invalid JSON — {ex.Message}";
        }

        if (patchArray.ValueKind != JsonValueKind.Array)
            return "Error: patches must be a JSON array.";

        var patchCount = patchArray.GetArrayLength();
        if (patchCount > 10)
            return $"Error: Too many operations ({patchCount}). Maximum is 10 per call. Split into multiple calls.";

        var results = new List<string>();
        var succeededPatches = new List<string>();
        int applied = 0;

        foreach (var patch in patchArray.EnumerateArray())
        {
            try
            {
                var op = patch.GetProperty("op").GetString()?.ToLowerInvariant()
                    ?? throw new ArgumentException("Patch must have an 'op' field.");

                switch (op)
                {
                    case "add":
                        ApplyAdd(patch, wpDoc, mainPart);
                        break;
                    case "replace":
                        ApplyReplace(patch, wpDoc, mainPart);
                        break;
                    case "remove":
                        ApplyRemove(patch, wpDoc);
                        break;
                    case "move":
                        ApplyMove(patch, wpDoc);
                        break;
                    case "copy":
                        ApplyCopy(patch, wpDoc);
                        break;
                    case "replace_text":
                        ApplyReplaceText(patch, wpDoc);
                        break;
                    case "remove_column":
                        ApplyRemoveColumn(patch, wpDoc);
                        break;
                    default:
                        results.Add($"Unknown operation: '{op}'");
                        continue;
                }

                succeededPatches.Add(patch.GetRawText());
                applied++;
            }
            catch (Exception ex)
            {
                var pathStr = patch.TryGetProperty("path", out var p) ? p.GetString() : "(no path)";
                results.Add($"Error at '{pathStr}': {ex.Message}");
            }
        }

        // Append only successful patches to WAL for replay fidelity
        if (succeededPatches.Count > 0)
        {
            try
            {
                var walPatches = $"[{string.Join(",", succeededPatches)}]";
                sessions.AppendWal(doc_id, walPatches);
            }
            catch { /* persistence is best-effort */ }
        }

        if (results.Count > 0)
            return $"Applied {applied}/{patchArray.GetArrayLength()} patches.\n" +
                   string.Join("\n", results);

        return $"Applied {applied} patch(es) successfully.";
    }

    internal static void ReplayAdd(JsonElement patch, WordprocessingDocument wpDoc, MainDocumentPart mainPart) => ApplyAdd(patch, wpDoc, mainPart);
    internal static void ReplayReplace(JsonElement patch, WordprocessingDocument wpDoc, MainDocumentPart mainPart) => ApplyReplace(patch, wpDoc, mainPart);
    internal static void ReplayRemove(JsonElement patch, WordprocessingDocument wpDoc) => ApplyRemove(patch, wpDoc);
    internal static void ReplayMove(JsonElement patch, WordprocessingDocument wpDoc) => ApplyMove(patch, wpDoc);
    internal static void ReplayCopy(JsonElement patch, WordprocessingDocument wpDoc) => ApplyCopy(patch, wpDoc);
    internal static void ReplayReplaceText(JsonElement patch, WordprocessingDocument wpDoc) => ApplyReplaceText(patch, wpDoc);
    internal static void ReplayRemoveColumn(JsonElement patch, WordprocessingDocument wpDoc) => ApplyRemoveColumn(patch, wpDoc);

    private static void ApplyAdd(JsonElement patch, WordprocessingDocument wpDoc, MainDocumentPart mainPart)
    {
        var pathStr = patch.GetProperty("path").GetString()
            ?? throw new ArgumentException("Add patch must have a 'path' field.");
        var value = patch.GetProperty("value");

        var path = DocxPath.Parse(pathStr);

        if (path.IsChildrenPath)
        {
            var (parent, index) = PathResolver.ResolveForInsert(path, wpDoc);

            if (value.TryGetProperty("type", out var typeEl) && typeEl.GetString() == "list")
            {
                var items = ElementFactory.CreateListItems(value);
                for (int i = items.Count - 1; i >= 0; i--)
                {
                    parent.InsertChildAt(items[i], index);
                }
            }
            else
            {
                var element = ElementFactory.CreateFromJson(value, mainPart);
                parent.InsertChildAt(element, index);
            }
        }
        else
        {
            var parentPath = new DocxPath(path.Segments.ToList());
            var parents = PathResolver.Resolve(parentPath, wpDoc);

            if (parents.Count != 1)
                throw new InvalidOperationException("Add path must resolve to exactly one parent.");

            var parent = parents[0];

            if (value.TryGetProperty("type", out var t) && t.GetString() == "list")
            {
                var items = ElementFactory.CreateListItems(value);
                foreach (var item in items)
                    parent.AppendChild(item);
            }
            else
            {
                var element = ElementFactory.CreateFromJson(value, mainPart);
                parent.AppendChild(element);
            }
        }
    }

    private static void ApplyReplace(JsonElement patch, WordprocessingDocument wpDoc, MainDocumentPart mainPart)
    {
        var pathStr = patch.GetProperty("path").GetString()
            ?? throw new ArgumentException("Replace patch must have a 'path' field.");
        var value = patch.GetProperty("value");

        var path = DocxPath.Parse(pathStr);
        var targets = PathResolver.Resolve(path, wpDoc);

        if (targets.Count == 0)
            throw new InvalidOperationException($"No elements found at path '{pathStr}'.");

        if (path.Leaf is StyleSegment)
        {
            foreach (var target in targets)
            {
                if (target is ParagraphProperties)
                {
                    var newProps = ElementFactory.CreateParagraphProperties(value);
                    target.Parent?.ReplaceChild(newProps, target);
                }
                else if (target is RunProperties)
                {
                    var newProps = ElementFactory.CreateRunProperties(value);
                    target.Parent?.ReplaceChild(newProps, target);
                }
                else if (target is TableProperties)
                {
                    var newProps = ElementFactory.CreateTableProperties(value);
                    target.Parent?.ReplaceChild(newProps, target);
                }
            }
            return;
        }

        foreach (var target in targets)
        {
            var parent = target.Parent
                ?? throw new InvalidOperationException("Target element has no parent.");

            var newElement = ElementFactory.CreateFromJson(value, mainPart);
            parent.ReplaceChild(newElement, target);
        }
    }

    private static void ApplyRemove(JsonElement patch, WordprocessingDocument wpDoc)
    {
        var pathStr = patch.GetProperty("path").GetString()
            ?? throw new ArgumentException("Remove patch must have a 'path' field.");

        var path = DocxPath.Parse(pathStr);
        var targets = PathResolver.Resolve(path, wpDoc);

        if (targets.Count == 0)
            throw new InvalidOperationException($"No elements found at path '{pathStr}'.");

        foreach (var target in targets)
        {
            target.Parent?.RemoveChild(target);
        }
    }

    private static void ApplyMove(JsonElement patch, WordprocessingDocument wpDoc)
    {
        var fromStr = patch.GetProperty("from").GetString()
            ?? throw new ArgumentException("Move patch must have a 'from' field.");
        var toStr = patch.GetProperty("path").GetString()
            ?? throw new ArgumentException("Move patch must have a 'path' field.");

        var fromPath = DocxPath.Parse(fromStr);
        var sources = PathResolver.Resolve(fromPath, wpDoc);
        if (sources.Count != 1)
            throw new InvalidOperationException("Move source must resolve to exactly one element.");

        var source = sources[0];
        source.Parent?.RemoveChild(source);

        var toPath = DocxPath.Parse(toStr);
        if (toPath.IsChildrenPath)
        {
            var (parent, index) = PathResolver.ResolveForInsert(toPath, wpDoc);
            parent.InsertChildAt(source, index);
        }
        else
        {
            var targets = PathResolver.Resolve(toPath, wpDoc);
            if (targets.Count != 1)
                throw new InvalidOperationException("Move target must resolve to exactly one location.");

            var target = targets[0];
            target.Parent?.InsertAfter(source, target);
        }
    }

    private static void ApplyCopy(JsonElement patch, WordprocessingDocument wpDoc)
    {
        var fromStr = patch.GetProperty("from").GetString()
            ?? throw new ArgumentException("Copy patch must have a 'from' field.");
        var toStr = patch.GetProperty("path").GetString()
            ?? throw new ArgumentException("Copy patch must have a 'path' field.");

        var fromPath = DocxPath.Parse(fromStr);
        var sources = PathResolver.Resolve(fromPath, wpDoc);
        if (sources.Count != 1)
            throw new InvalidOperationException("Copy source must resolve to exactly one element.");

        var clone = sources[0].CloneNode(true);

        var toPath = DocxPath.Parse(toStr);
        if (toPath.IsChildrenPath)
        {
            var (parent, index) = PathResolver.ResolveForInsert(toPath, wpDoc);
            parent.InsertChildAt(clone, index);
        }
        else
        {
            var targets = PathResolver.Resolve(toPath, wpDoc);
            if (targets.Count != 1)
                throw new InvalidOperationException("Copy target must resolve to exactly one location.");

            var target = targets[0];
            target.Parent?.InsertAfter(clone, target);
        }
    }

    /// <summary>
    /// Find and replace text within runs, preserving all run-level formatting.
    /// Works on paragraphs, headings, table cells, or any element containing runs.
    /// </summary>
    private static void ApplyReplaceText(JsonElement patch, WordprocessingDocument wpDoc)
    {
        var pathStr = patch.GetProperty("path").GetString()
            ?? throw new ArgumentException("replace_text must have a 'path' field.");
        var find = patch.GetProperty("find").GetString()
            ?? throw new ArgumentException("replace_text must have a 'find' field.");
        var replace = patch.GetProperty("replace").GetString()
            ?? throw new ArgumentException("replace_text must have a 'replace' field.");

        var path = DocxPath.Parse(pathStr);
        var targets = PathResolver.Resolve(path, wpDoc);

        if (targets.Count == 0)
            throw new InvalidOperationException($"No elements found at path '{pathStr}'.");

        foreach (var target in targets)
        {
            ReplaceTextInElement(target, find, replace);
        }
    }

    /// <summary>
    /// Replace text within an element's runs, preserving formatting.
    /// Handles text that spans across multiple runs by performing per-run replacement
    /// for simple cases, and cross-run replacement for multi-run spans.
    /// </summary>
    private static void ReplaceTextInElement(OpenXmlElement element, string find, string replace)
    {
        // Collect all paragraphs (direct or nested)
        var paragraphs = element is Paragraph p
            ? new List<Paragraph> { p }
            : element.Descendants<Paragraph>().ToList();

        foreach (var para in paragraphs)
        {
            var runs = para.Elements<Run>().ToList();
            if (runs.Count == 0) continue;

            // Try simple per-run replacement first
            bool found = false;
            foreach (var run in runs)
            {
                var textElem = run.GetFirstChild<Text>();
                if (textElem is null) continue;

                var text = textElem.Text;
                if (text.Contains(find, StringComparison.Ordinal))
                {
                    textElem.Text = text.Replace(find, replace, StringComparison.Ordinal);
                    found = true;
                }
            }

            if (found) continue;

            // Cross-run replacement: concatenate all run texts, find the match,
            // then adjust the runs that contain the match
            var allText = string.Concat(runs.Select(r => r.InnerText));
            var matchIdx = allText.IndexOf(find, StringComparison.Ordinal);
            if (matchIdx < 0) continue;

            // Map character positions to runs
            int pos = 0;
            foreach (var run in runs)
            {
                var textElem = run.GetFirstChild<Text>();
                if (textElem is null)
                {
                    // Tab or break: count as 1 char (\t or empty)
                    var runText = run.InnerText;
                    pos += runText.Length;
                    continue;
                }

                var runStart = pos;
                var runEnd = pos + textElem.Text.Length;

                // Check if this run overlaps with the find range
                var findEnd = matchIdx + find.Length;

                if (runEnd <= matchIdx || runStart >= findEnd)
                {
                    // No overlap
                    pos = runEnd;
                    continue;
                }

                // This run overlaps with the search text
                var overlapStart = Math.Max(matchIdx, runStart) - runStart;
                var overlapEnd = Math.Min(findEnd, runEnd) - runStart;

                var before = textElem.Text[..overlapStart];
                var after = textElem.Text[overlapEnd..];

                // First overlapping run gets the replacement text
                if (runStart <= matchIdx)
                {
                    textElem.Text = before + replace + after;
                    textElem.Space = SpaceProcessingModeValues.Preserve;
                }
                else
                {
                    // Subsequent overlapping runs: remove the overlapping portion
                    textElem.Text = after;
                    textElem.Space = SpaceProcessingModeValues.Preserve;
                }

                pos = runEnd;
            }
        }
    }

    /// <summary>
    /// Remove a column from a table by index (0-based).
    /// Removes the cell at the given column index from every row.
    /// </summary>
    private static void ApplyRemoveColumn(JsonElement patch, WordprocessingDocument wpDoc)
    {
        var pathStr = patch.GetProperty("path").GetString()
            ?? throw new ArgumentException("remove_column must have a 'path' field.");
        var column = patch.GetProperty("column").GetInt32();

        var path = DocxPath.Parse(pathStr);
        var targets = PathResolver.Resolve(path, wpDoc);

        if (targets.Count == 0)
            throw new InvalidOperationException($"No elements found at path '{pathStr}'.");

        foreach (var target in targets)
        {
            if (target is not Table table)
                throw new InvalidOperationException("remove_column target must be a table.");

            foreach (var row in table.Elements<TableRow>())
            {
                var cells = row.Elements<TableCell>().ToList();
                if (column >= 0 && column < cells.Count)
                {
                    row.RemoveChild(cells[column]);
                }
            }

            // Update grid columns if present
            var grid = table.GetFirstChild<TableGrid>();
            if (grid is not null)
            {
                var gridCols = grid.Elements<GridColumn>().ToList();
                if (column >= 0 && column < gridCols.Count)
                {
                    grid.RemoveChild(gridCols[column]);
                }
            }
        }
    }
}
