using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace DocxMcp.Layout;

/// <summary>
/// `docx-cli dump <file.docx> [-o out.json]`: a stateless, read-only dump of a .docx
/// as JSON — what a mechanical parser needs to rebuild the document's
/// structure without interpreting it: blocks in document order
/// (paragraphs, tables, text boxes), runs with their EFFECTIVE formatting
/// (document defaults, then the paragraph style chain, the character style
/// chain, then direct formatting), numbering, headers and footers, images
/// (base64), page setup and theme colours. Sizes in points, colours #rrggbb.
/// </summary>
public static class TechnicalDump
{
    public const int Schema = 1;

    sealed record RunFormat(string? Font, double? Size, bool Bold, bool Italic, bool Underline, string? Color)
    {
        public static readonly RunFormat Empty = new(null, null, false, false, false, null);

        public RunFormat Over(RunFormat? below) => below is null
            ? this
            : new RunFormat(Font ?? below.Font, Size ?? below.Size, Bold || below.Bold, Italic || below.Italic,
                Underline || below.Underline, Color ?? below.Color);
    }

    sealed class Context(WordprocessingDocument doc)
    {
        public readonly WordprocessingDocument Doc = doc;
        public readonly W.Styles? Styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        public readonly Dictionary<string, string> ThemeColors = ReadThemeColors(doc);
        public readonly Dictionary<string, (string Mime, byte[] Bytes)> Images = new();
        public readonly HashSet<string> SeenTextBoxes = new();
    }

    /// The JSON dump of a .docx (UTF-8).
    public static byte[] Dump(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var ctx = new Context(doc);

        using var buffer = new MemoryStream();
        using (var w = new Utf8JsonWriter(buffer, new JsonWriterOptions { Indented = false }))
        {
            w.WriteStartObject();
            w.WriteNumber("schema", Schema);
            WritePage(w, doc);
            w.WriteStartObject("themeColors");
            foreach (var (k, v) in ctx.ThemeColors) w.WriteString(k, v);
            w.WriteEndObject();

            var body = doc.MainDocumentPart?.Document?.Body;
            w.WriteStartArray("body");
            if (body is not null) WriteBlocks(w, ctx, body.ChildElements, doc.MainDocumentPart!);
            w.WriteEndArray();

            w.WriteStartArray("headers");
            foreach (var part in doc.MainDocumentPart?.HeaderParts ?? [])
                if (part.Header is not null) WriteBlocks(w, ctx, part.Header.ChildElements, part);
            w.WriteEndArray();
            w.WriteStartArray("footers");
            foreach (var part in doc.MainDocumentPart?.FooterParts ?? [])
                if (part.Footer is not null) WriteBlocks(w, ctx, part.Footer.ChildElements, part);
            w.WriteEndArray();

            w.WriteStartArray("images");
            foreach (var (id, (mime, data)) in ctx.Images)
            {
                w.WriteStartObject();
                w.WriteString("id", id);
                w.WriteString("mime", mime);
                w.WriteString("base64", Convert.ToBase64String(data));
                w.WriteEndObject();
            }
            w.WriteEndArray();
            w.WriteEndObject();
        }
        return buffer.ToArray();
    }

    static double Pt(uint? twips) => twips is null ? 0 : twips.Value / 20.0;

    static void WritePage(Utf8JsonWriter w, WordprocessingDocument doc)
    {
        var sect = doc.MainDocumentPart?.Document?.Body?.GetFirstChild<W.SectionProperties>();
        var size = sect?.GetFirstChild<W.PageSize>();
        var margin = sect?.GetFirstChild<W.PageMargin>();
        var cols = sect?.GetFirstChild<W.Columns>();
        w.WriteStartObject("page");
        w.WriteNumber("width", Pt(size?.Width?.Value ?? 11906));
        w.WriteNumber("height", Pt(size?.Height?.Value ?? 16838));
        w.WriteNumber("marginTop", (margin?.Top?.Value ?? 1440) / 20.0);
        w.WriteNumber("marginBottom", (margin?.Bottom?.Value ?? 1440) / 20.0);
        w.WriteNumber("marginLeft", Pt(margin?.Left?.Value ?? 1440));
        w.WriteNumber("marginRight", Pt(margin?.Right?.Value ?? 1440));
        w.WriteNumber("columns", cols?.ColumnCount?.Value ?? 1);
        w.WriteEndObject();
    }

    static Dictionary<string, string> ReadThemeColors(WordprocessingDocument doc)
    {
        var colors = new Dictionary<string, string>();
        var scheme = doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        if (scheme is null) return colors;
        foreach (var child in scheme.ChildElements)
        {
            var rgb = child.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value
                ?? child.GetFirstChild<A.SystemColor>()?.LastColor?.Value;
            if (rgb is not null) colors[child.LocalName] = "#" + rgb.ToLowerInvariant();
        }
        return colors;
    }

    // --- formatting resolution ------------------------------------------------

    static RunFormat FromRunProps(OpenXmlElement? rPr, Context ctx)
    {
        if (rPr is null) return RunFormat.Empty;
        var fonts = rPr.GetFirstChild<W.RunFonts>();
        var size = rPr.GetFirstChild<W.FontSize>()?.Val?.Value;
        var color = rPr.GetFirstChild<W.Color>();
        string? hex = null;
        if (color?.ThemeColor is { HasValue: true } theme && ctx.ThemeColors.TryGetValue(ThemeKey(theme.InnerText!), out var themed))
            hex = themed;
        else if (color?.Val?.Value is { } val && val != "auto")
            hex = "#" + val.ToLowerInvariant();
        static bool On(W.OnOffType? t) => t is not null && (t.Val is null || t.Val.Value);
        return new RunFormat(
            fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value ?? fonts?.ComplexScript?.Value,
            size is null ? null : double.Parse(size) / 2.0,
            On(rPr.GetFirstChild<W.Bold>()),
            On(rPr.GetFirstChild<W.Italic>()),
            rPr.GetFirstChild<W.Underline>()?.Val?.Value is { } u && u != W.UnderlineValues.None,
            hex);
    }

    /// Theme colour names in w:color (accent1, text1…) and in the theme part (accent1, dk1…).
    static string ThemeKey(string name) => name.ToLowerInvariant() switch
    {
        "text1" or "dark1" => "dk1",
        "text2" or "dark2" => "dk2",
        "background1" or "light1" => "lt1",
        "background2" or "light2" => "lt2",
        var n => n,
    };

    static W.Style? Style(Context ctx, string? id) =>
        id is null ? null : ctx.Styles?.Elements<W.Style>().FirstOrDefault(s => s.StyleId?.Value == id);

    /// The run formatting a style chain gives (the style over its bases).
    static RunFormat StyleChain(Context ctx, string? id, int depth = 0)
    {
        var style = Style(ctx, id);
        if (style is null || depth > 16) return RunFormat.Empty;
        var own = FromRunProps(style.StyleRunProperties, ctx);
        return own.Over(StyleChain(ctx, style.BasedOn?.Val?.Value, depth + 1));
    }

    static RunFormat Defaults(Context ctx) =>
        FromRunProps(ctx.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle, ctx);

    static string? DefaultParagraphStyle(Context ctx) =>
        ctx.Styles?.Elements<W.Style>()
            .FirstOrDefault(s => s.Type?.Value == W.StyleValues.Paragraph && s.Default?.Value == true)?.StyleId?.Value;

    // --- blocks -----------------------------------------------------------------

    static void WriteBlocks(Utf8JsonWriter w, Context ctx, IEnumerable<OpenXmlElement> elements, OpenXmlPart part)
    {
        foreach (var element in elements)
        {
            switch (element)
            {
                case W.Paragraph p:
                    WriteParagraph(w, ctx, p, part);
                    break;
                case W.Table t:
                    WriteTable(w, ctx, t, part);
                    break;
                case W.SdtBlock sdt when sdt.SdtContentBlock is not null:
                    WriteBlocks(w, ctx, sdt.SdtContentBlock.ChildElements, part);
                    break;
            }
        }
    }

    static void WriteTable(Utf8JsonWriter w, Context ctx, W.Table table, OpenXmlPart part)
    {
        w.WriteStartObject();
        w.WriteString("type", "table");
        w.WriteStartArray("rows");
        foreach (var row in table.Elements<W.TableRow>())
        {
            w.WriteStartArray();
            foreach (var cell in row.Elements<W.TableCell>())
            {
                w.WriteStartObject();
                var fill = cell.TableCellProperties?.Shading?.Fill?.Value;
                if (fill is not null && fill != "auto") w.WriteString("shading", "#" + fill.ToLowerInvariant());
                var width = cell.TableCellProperties?.TableCellWidth?.Width?.Value;
                if (width is not null && cell.TableCellProperties?.TableCellWidth?.Type?.Value == W.TableWidthUnitValues.Dxa)
                    w.WriteNumber("width", double.Parse(width) / 20.0);
                w.WriteStartArray("blocks");
                WriteBlocks(w, ctx, cell.ChildElements, part);
                w.WriteEndArray();
                w.WriteEndObject();
            }
            w.WriteEndArray();
        }
        w.WriteEndArray();
        w.WriteEndObject();
    }

    static void WriteParagraph(Utf8JsonWriter w, Context ctx, W.Paragraph p, OpenXmlPart part)
    {
        var pPr = p.ParagraphProperties;
        var styleId = pPr?.ParagraphStyleId?.Val?.Value ?? DefaultParagraphStyle(ctx);
        var style = Style(ctx, styleId);
        var baseFormat = StyleChain(ctx, styleId).Over(Defaults(ctx));

        w.WriteStartObject();
        w.WriteString("type", "paragraph");
        if (styleId is not null) w.WriteString("style", styleId);
        if (style?.StyleName?.Val?.Value is { } name) w.WriteString("styleName", name);
        var outline = pPr?.OutlineLevel?.Val?.Value ?? style?.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
        if (outline is not null) w.WriteNumber("outlineLevel", outline.Value);
        var numPr = pPr?.NumberingProperties ?? style?.StyleParagraphProperties?.NumberingProperties;
        if (numPr?.NumberingId?.Val?.Value is { } numId && numId != 0)
        {
            w.WriteStartObject("list");
            w.WriteNumber("id", numId);
            w.WriteNumber("level", numPr.NumberingLevelReference?.Val?.Value ?? 0);
            w.WriteEndObject();
        }
        if (pPr?.Justification?.Val is { HasValue: true } jc) w.WriteString("align", jc.InnerText);
        var shading = pPr?.Shading?.Fill?.Value;
        if (shading is not null && shading != "auto") w.WriteString("shading", "#" + shading.ToLowerInvariant());

        w.WriteStartArray("runs");
        var textBoxes = new List<OpenXmlElement>();
        var images = new List<string>();
        // Runs inside a text box anchored in this paragraph belong to that
        // text box (dumped after it), not to the paragraph.
        var ownBox = p.Ancestors<W.TextBoxContent>().FirstOrDefault();
        foreach (var run in p.Descendants<W.Run>())
        {
            if (run.Ancestors<W.TextBoxContent>().FirstOrDefault() != ownBox) continue;
            var text = new StringBuilder();
            foreach (var child in run.ChildElements)
            {
                switch (child)
                {
                    case W.Text t: text.Append(t.Text); break;
                    case W.TabChar: text.Append('\t'); break;
                    case W.Break: text.Append('\n'); break;
                }
            }
            foreach (var blip in run.Descendants<A.Blip>())
                if (blip.Embed?.Value is { } relId && Image(ctx, part, relId) is { } imageId) images.Add(imageId);
            foreach (var box in run.Descendants<W.TextBoxContent>())
                if (box.Ancestors<W.TextBoxContent>().FirstOrDefault() == ownBox) textBoxes.Add(box);
            if (text.Length == 0) continue;
            var characterStyle = run.RunProperties?.RunStyle?.Val?.Value;
            var format = FromRunProps(run.RunProperties, ctx)
                .Over(StyleChain(ctx, characterStyle))
                .Over(baseFormat);
            w.WriteStartObject();
            w.WriteString("text", text.ToString());
            if (format.Font is not null) w.WriteString("font", format.Font);
            if (format.Size is not null) w.WriteNumber("size", format.Size.Value);
            w.WriteBoolean("bold", format.Bold);
            w.WriteBoolean("italic", format.Italic);
            if (format.Underline) w.WriteBoolean("underline", true);
            w.WriteString("color", format.Color ?? "#000000");
            w.WriteEndObject();
        }
        w.WriteEndArray();
        if (images.Count > 0)
        {
            w.WriteStartArray("images");
            foreach (var id in images) w.WriteStringValue(id);
            w.WriteEndArray();
        }
        w.WriteEndObject();

        // Text boxes (floating zones of graphic CVs) follow their anchor.
        foreach (var box in textBoxes)
        {
            if (!ctx.SeenTextBoxes.Add(box.OuterXml)) continue;   // mc:Choice and mc:Fallback hold the same box
            w.WriteStartObject();
            w.WriteString("type", "textbox");
            w.WriteStartArray("blocks");
            WriteBlocks(w, ctx, box.ChildElements, part);
            w.WriteEndArray();
            w.WriteEndObject();
        }
    }

    static string? Image(Context ctx, OpenXmlPart part, string relId)
    {
        if (part.GetPartById(relId) is not ImagePart image) return null;
        var id = image.Uri.ToString();
        if (!ctx.Images.ContainsKey(id))
        {
            using var s = image.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            ctx.Images[id] = (image.ContentType, ms.ToArray());
        }
        return id;
    }
}
