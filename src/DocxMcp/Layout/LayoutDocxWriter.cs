using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace DocxMcp.Layout;

/// <summary>Options of <see cref="LayoutDocxWriter"/>.</summary>
public sealed record LayoutDocxOptions
{
    /// <summary>Vertical placement model (see <see cref="BaselineModel"/>).</summary>
    public BaselineModel Baseline { get; init; } = BaselineModel.Default;

    /// <summary>
    /// Horizontal slack (pt) added to text boxes on the side(s) toward which a line may grow,
    /// so a line that Word measures marginally wider than Typst does not wrap.
    /// Paragraph indents compensate, so the text itself stays at the layout position.
    /// </summary>
    public double HorizontalSlack { get; init; } = 24;

    /// <summary>Normalize zip timestamps so the same input yields the same bytes.</summary>
    public bool Deterministic { get; init; } = true;
}

/// <summary>
/// Writes a brand new .docx from an absolute <see cref="LayoutDocument"/>: one Word section per
/// layout page (exact page size, zero margins), a tiny anchor paragraph per page, and every
/// layout item as an anchored DrawingML object positioned relative to the page, in paint order.
/// </summary>
public static class LayoutDocxWriter
{
    public const long EmuPerPoint = 12700;

    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private const string WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    private const string WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
    private const string W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string HyperlinkRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    public static void Write(LayoutDocument layout, Stream output) => Write(layout, output, new LayoutDocxOptions());

    public static void Write(LayoutDocument layout, Stream output, LayoutDocxOptions options)
    {
        ArgumentNullException.ThrowIfNull(layout);
        ArgumentNullException.ThrowIfNull(output);

        using var package = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(package, WordprocessingDocumentType.Document))
        {
            // Fixed relationship id (AddMainDocumentPart() would generate a random one).
            var main = doc.AddNewPart<MainDocumentPart>(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1");
            var ctx = new WriteContext(main, options);

            Feed(main.AddNewPart<StyleDefinitionsPart>("rIdStyles"), StylesXml);
            Feed(main.AddNewPart<DocumentSettingsPart>("rIdSettings"), SettingsXml);

            var body = BuildDocumentXml(layout, ctx);
            Feed(main, body);
        }

        package.Position = 0;
        if (options.Deterministic)
            NormalizeZip(package, output);
        else
            package.CopyTo(output);
    }

    /// <summary>Convenience: parse layout JSON and write the .docx to <paramref name="outputPath"/>.</summary>
    public static void WriteFile(string layoutJson, string outputPath, LayoutDocxOptions? options = null)
    {
        var layout = LayoutParser.Parse(layoutJson);
        using var fs = File.Create(outputPath);
        Write(layout, fs, options ?? new LayoutDocxOptions());
    }

    // ------------------------------------------------------------------ document

    private sealed class WriteContext(MainDocumentPart main, LayoutDocxOptions options)
    {
        public MainDocumentPart Main { get; } = main;
        public LayoutDocxOptions Options { get; } = options;
        public uint NextDocPrId = 1;
        public uint NextZ = 1;
        public int NextImage = 1;
        public readonly Dictionary<string, string> LinkIds = new(StringComparer.Ordinal);
    }

    private static byte[] BuildDocumentXml(LayoutDocument layout, WriteContext ctx)
    {
        var ms = new MemoryStream();
        using (var x = XmlWriter.Create(ms, new XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
        {
            x.WriteStartDocument(true);
            x.WriteStartElement("w", "document", W);
            x.WriteAttributeString("xmlns", "w", null, W);
            x.WriteAttributeString("xmlns", "r", null, R);
            x.WriteAttributeString("xmlns", "wp", null, WP);
            x.WriteAttributeString("xmlns", "a", null, A);
            x.WriteAttributeString("xmlns", "pic", null, PIC);
            x.WriteAttributeString("xmlns", "wps", null, WPS);
            x.WriteAttributeString("xmlns", "wp14", null, WP14);
            x.WriteAttributeString("xmlns", "w14", null, W14);
            x.WriteAttributeString("xmlns", "mc", null, MC);
            x.WriteAttributeString("mc", "Ignorable", MC, "w14 wp14");
            x.WriteStartElement("w", "body", W);

            for (var i = 0; i < layout.Pages.Count; i++)
            {
                var page = layout.Pages[i];
                var last = i == layout.Pages.Count - 1;
                WritePageParagraph(x, page, ctx, sectionBreak: !last);
                if (last)
                    WriteSectPr(x, page);
            }

            x.WriteEndElement(); // body
            x.WriteEndElement(); // document
        }
        return ms.ToArray();
    }

    /// <summary>The per-page anchor paragraph: 1pt exact line, no spacing, carries all drawings.</summary>
    private static void WritePageParagraph(XmlWriter x, LayoutPage page, WriteContext ctx, bool sectionBreak)
    {
        x.WriteStartElement("w", "p", W);
        x.WriteStartElement("w", "pPr", W);
        WriteSpacing(x, 20);
        x.WriteStartElement("w", "rPr", W);
        WriteVal(x, "sz", "2");
        WriteVal(x, "szCs", "2");
        x.WriteEndElement();
        if (sectionBreak)
            WriteSectPr(x, page);
        x.WriteEndElement(); // pPr

        foreach (var item in page.Items)
        {
            x.WriteStartElement("w", "r", W);
            x.WriteStartElement("w", "rPr", W);
            WriteVal(x, "sz", "2");
            WriteVal(x, "szCs", "2");
            x.WriteEndElement();
            x.WriteStartElement("w", "drawing", W);
            switch (item)
            {
                case LayoutText t: WriteTextBox(x, t, ctx); break;
                case LayoutRect r: WriteRect(x, r, ctx); break;
                case LayoutLineShape l: WriteLine(x, l, ctx); break;
                case LayoutImage img: WriteImage(x, img, ctx); break;
            }
            x.WriteEndElement(); // drawing
            x.WriteEndElement(); // r
        }

        x.WriteEndElement(); // p
    }

    private static void WriteSectPr(XmlWriter x, LayoutPage page)
    {
        x.WriteStartElement("w", "sectPr", W);
        WriteVal(x, "type", "nextPage");
        x.WriteStartElement("w", "pgSz", W);
        x.WriteAttributeString("w", "w", W, Twips(page.Width).ToString(CultureInfo.InvariantCulture));
        x.WriteAttributeString("w", "h", W, Twips(page.Height).ToString(CultureInfo.InvariantCulture));
        if (page.Width > page.Height)
            x.WriteAttributeString("w", "orient", W, "landscape");
        x.WriteEndElement();
        x.WriteStartElement("w", "pgMar", W);
        foreach (var side in new[] { "top", "right", "bottom", "left", "header", "footer", "gutter" })
            x.WriteAttributeString("w", side, W, "0");
        x.WriteEndElement();
        x.WriteStartElement("w", "cols", W);
        x.WriteAttributeString("w", "space", W, "0");
        x.WriteEndElement();
        x.WriteEndElement();
    }

    // ------------------------------------------------------------------ anchors

    private static void BeginAnchor(XmlWriter x, WriteContext ctx, double xPt, double yPt, double wPt, double hPt,
        string name, double effectPt = 0)
    {
        x.WriteStartElement("wp", "anchor", WP);
        foreach (var d in new[] { "distT", "distB", "distL", "distR" })
            x.WriteAttributeString(d, "0");
        x.WriteAttributeString("simplePos", "0");
        x.WriteAttributeString("relativeHeight", (ctx.NextZ++).ToString(CultureInfo.InvariantCulture));
        x.WriteAttributeString("behindDoc", "0");
        x.WriteAttributeString("locked", "0");
        x.WriteAttributeString("layoutInCell", "1");
        x.WriteAttributeString("allowOverlap", "1");

        x.WriteStartElement("wp", "simplePos", WP);
        x.WriteAttributeString("x", "0");
        x.WriteAttributeString("y", "0");
        x.WriteEndElement();

        x.WriteStartElement("wp", "positionH", WP);
        x.WriteAttributeString("relativeFrom", "page");
        x.WriteElementString("wp", "posOffset", WP, Emu(xPt).ToString(CultureInfo.InvariantCulture));
        x.WriteEndElement();
        x.WriteStartElement("wp", "positionV", WP);
        x.WriteAttributeString("relativeFrom", "page");
        x.WriteElementString("wp", "posOffset", WP, Emu(yPt).ToString(CultureInfo.InvariantCulture));
        x.WriteEndElement();

        x.WriteStartElement("wp", "extent", WP);
        x.WriteAttributeString("cx", Emu(Math.Max(0, wPt)).ToString(CultureInfo.InvariantCulture));
        x.WriteAttributeString("cy", Emu(Math.Max(0, hPt)).ToString(CultureInfo.InvariantCulture));
        x.WriteEndElement();

        var e = Emu(Math.Max(0, effectPt)).ToString(CultureInfo.InvariantCulture);
        x.WriteStartElement("wp", "effectExtent", WP);
        x.WriteAttributeString("l", e);
        x.WriteAttributeString("t", e);
        x.WriteAttributeString("r", e);
        x.WriteAttributeString("b", e);
        x.WriteEndElement();

        x.WriteElementString("wp", "wrapNone", WP, null);

        var id = ctx.NextDocPrId++;
        x.WriteStartElement("wp", "docPr", WP);
        x.WriteAttributeString("id", id.ToString(CultureInfo.InvariantCulture));
        x.WriteAttributeString("name", $"{name} {id}");
        x.WriteEndElement();
        x.WriteElementString("wp", "cNvGraphicFramePr", WP, null);

        x.WriteStartElement("a", "graphic", A);
        x.WriteStartElement("a", "graphicData", A);
    }

    private static void EndAnchor(XmlWriter x)
    {
        x.WriteEndElement(); // graphicData
        x.WriteEndElement(); // graphic
        x.WriteEndElement(); // anchor
    }

    private static void WriteXfrm(XmlWriter x, double wPt, double hPt, bool flipH = false, bool flipV = false)
    {
        x.WriteStartElement("a", "xfrm", A);
        if (flipH) x.WriteAttributeString("flipH", "1");
        if (flipV) x.WriteAttributeString("flipV", "1");
        x.WriteStartElement("a", "off", A);
        x.WriteAttributeString("x", "0");
        x.WriteAttributeString("y", "0");
        x.WriteEndElement();
        x.WriteStartElement("a", "ext", A);
        x.WriteAttributeString("cx", Emu(Math.Max(0, wPt)).ToString(CultureInfo.InvariantCulture));
        x.WriteAttributeString("cy", Emu(Math.Max(0, hPt)).ToString(CultureInfo.InvariantCulture));
        x.WriteEndElement();
        x.WriteEndElement();
    }

    private static void WritePrstGeom(XmlWriter x, string prst)
    {
        x.WriteStartElement("a", "prstGeom", A);
        x.WriteAttributeString("prst", prst);
        x.WriteElementString("a", "avLst", A, null);
        x.WriteEndElement();
    }

    private static void WriteSolidFill(XmlWriter x, string color)
    {
        x.WriteStartElement("a", "solidFill", A);
        x.WriteStartElement("a", "srgbClr", A);
        x.WriteAttributeString("val", color);
        x.WriteEndElement();
        x.WriteEndElement();
    }

    private static void WriteLn(XmlWriter x, LayoutStroke? stroke)
    {
        x.WriteStartElement("a", "ln", A);
        if (stroke is null || stroke.Width <= 0)
        {
            x.WriteElementString("a", "noFill", A, null);
        }
        else
        {
            x.WriteAttributeString("w", Emu(stroke.Width).ToString(CultureInfo.InvariantCulture));
            x.WriteAttributeString("cap", "flat");
            WriteSolidFill(x, stroke.Color);
            x.WriteStartElement("a", "miter", A);
            x.WriteAttributeString("lim", "800000");
            x.WriteEndElement();
        }
        x.WriteEndElement();
    }

    private static void BeginWsp(XmlWriter x, bool textBox)
    {
        x.WriteAttributeString("uri", WPS);
        x.WriteStartElement("wps", "wsp", WPS);
        x.WriteStartElement("wps", "cNvSpPr", WPS);
        if (textBox) x.WriteAttributeString("txBox", "1");
        x.WriteEndElement();
    }

    private static void WriteEmptyBodyPr(XmlWriter x)
    {
        x.WriteStartElement("wps", "bodyPr", WPS);
        x.WriteAttributeString("rot", "0");
        x.WriteAttributeString("vert", "horz");
        x.WriteAttributeString("wrap", "square");
        foreach (var ins in new[] { "lIns", "tIns", "rIns", "bIns" })
            x.WriteAttributeString(ins, "0");
        x.WriteAttributeString("anchor", "t");
        x.WriteAttributeString("anchorCtr", "0");
        x.WriteElementString("a", "noAutofit", A, null);
        x.WriteEndElement();
    }

    // ------------------------------------------------------------------ rect / line / image

    private static void WriteRect(XmlWriter x, LayoutRect r, WriteContext ctx)
    {
        var sw = r.Stroke?.Width ?? 0;
        BeginAnchor(x, ctx, r.X, r.Y, r.Width, r.Height, "Rectangle", sw / 2);
        BeginWsp(x, textBox: false);
        x.WriteStartElement("wps", "spPr", WPS);
        WriteXfrm(x, r.Width, r.Height);
        WritePrstGeom(x, "rect");
        if (r.Fill is null) x.WriteElementString("a", "noFill", A, null);
        else WriteSolidFill(x, r.Fill);
        WriteLn(x, r.Stroke);
        x.WriteEndElement(); // spPr
        WriteEmptyBodyPr(x);
        x.WriteEndElement(); // wsp
        EndAnchor(x);
    }

    private static void WriteLine(XmlWriter x, LayoutLineShape l, WriteContext ctx)
    {
        var left = Math.Min(l.X1, l.X2);
        var top = Math.Min(l.Y1, l.Y2);
        var w = Math.Abs(l.X2 - l.X1);
        var h = Math.Abs(l.Y2 - l.Y1);
        // The preset line runs from the top-left to the bottom-right corner of its box;
        // a line going up (or down-left) is that line mirrored horizontally.
        var flipH = (l.X2 - l.X1) * (l.Y2 - l.Y1) < 0;

        BeginAnchor(x, ctx, left, top, w, h, "Line", l.Stroke.Width / 2);
        BeginWsp(x, textBox: false);
        x.WriteStartElement("wps", "spPr", WPS);
        WriteXfrm(x, w, h, flipH: flipH);
        WritePrstGeom(x, "line");
        WriteLn(x, l.Stroke);
        x.WriteEndElement(); // spPr
        WriteEmptyBodyPr(x);
        x.WriteEndElement(); // wsp
        EndAnchor(x);
    }

    private static void WriteImage(XmlWriter x, LayoutImage img, WriteContext ctx)
    {
        var type = img.Format switch
        {
            "png" => ImagePartType.Png,
            "jpeg" or "jpg" => ImagePartType.Jpeg,
            "gif" => ImagePartType.Gif,
            "bmp" => ImagePartType.Bmp,
            "tiff" or "tif" => ImagePartType.Tiff,
            _ => (PartTypeInfo?)null,
        };
        if (type is null)
        {
            // SVG and other vector formats are not supported yet: skip, keep going.
            Console.Error.WriteLine($"[from-layout] warning: image format '{img.Format}' not supported, skipped");
            // Keep the run valid: an empty drawing is not allowed, so write an empty rectangle placeholder.
            WriteRect(x, new LayoutRect(img.X, img.Y, img.Width, img.Height, null, null), ctx);
            return;
        }

        var relId = $"rIdImg{ctx.NextImage++}";
        var part = ctx.Main.AddImagePart(type.Value, relId);
        using (var ms = new MemoryStream(img.Data))
            part.FeedData(ms);

        BeginAnchor(x, ctx, img.X, img.Y, img.Width, img.Height, "Picture");
        x.WriteAttributeString("uri", PIC);
        x.WriteStartElement("pic", "pic", PIC);
        x.WriteStartElement("pic", "nvPicPr", PIC);
        x.WriteStartElement("pic", "cNvPr", PIC);
        x.WriteAttributeString("id", "0");
        x.WriteAttributeString("name", relId);
        x.WriteEndElement();
        x.WriteStartElement("pic", "cNvPicPr", PIC);
        x.WriteStartElement("a", "picLocks", A);
        x.WriteAttributeString("noChangeAspect", "1");
        x.WriteEndElement();
        x.WriteEndElement();
        x.WriteEndElement(); // nvPicPr
        x.WriteStartElement("pic", "blipFill", PIC);
        x.WriteStartElement("a", "blip", A);
        x.WriteAttributeString("r", "embed", R, relId);
        x.WriteEndElement();
        x.WriteStartElement("a", "stretch", A);
        x.WriteElementString("a", "fillRect", A, null);
        x.WriteEndElement();
        x.WriteEndElement(); // blipFill
        x.WriteStartElement("pic", "spPr", PIC);
        WriteXfrm(x, img.Width, img.Height);
        WritePrstGeom(x, "rect");
        x.WriteEndElement();
        x.WriteEndElement(); // pic
        EndAnchor(x);
    }

    // ------------------------------------------------------------------ text

    /// <summary>
    /// Vertical plan of a text block: the text box top and each line's exact Word line height,
    /// chosen so that every line's baseline lands on the layout baseline (errors do not accumulate:
    /// each line is solved against its absolute target).
    /// </summary>
    public sealed record TextPlan(double BoxTop, IReadOnlyList<double> LineHeights)
    {
        public double ContentHeight => LineHeights.Sum();
    }

    public static TextPlan PlanText(LayoutText t, BaselineModel model)
    {
        var heights = new List<double>();
        if (t.Lines.Count == 0)
            return new TextPlan(t.Y, heights);

        static double SizeOf(LayoutLine l) => l.MaxSize > 0 ? l.MaxSize : 10;

        // First line: keep at least the font size so glyphs are not clipped by the exact height.
        var first = t.Lines[0];
        var s0 = SizeOf(first);
        var h0 = RoundTwips(Math.Max(first.Height > 0 ? first.Height : s0 * 1.2, s0));
        var boxTop = t.Y + first.Baseline - model.BaselineOffset(h0, s0);
        heights.Add(h0);

        var top = boxTop + h0;
        for (var k = 1; k < t.Lines.Count; k++)
        {
            var line = t.Lines[k];
            var s = SizeOf(line);
            var target = t.Y + line.Baseline - top;
            var h = model.SolveLineHeight(target, s) ?? 0.05;
            h = Math.Max(RoundTwips(h), 0.05);
            heights.Add(h);
            top += h;
        }
        return new TextPlan(boxTop, heights);
    }

    private static void WriteTextBox(XmlWriter x, LayoutText t, WriteContext ctx)
    {
        var plan = PlanText(t, ctx.Options.Baseline);
        var slack = Math.Max(0, ctx.Options.HorizontalSlack);
        var lastSize = t.Lines.Count > 0 ? t.Lines[^1].MaxSize : 0;
        // Extra room at the bottom: Word hides text overflowing a text box.
        var boxHeight = plan.ContentHeight + Math.Max(lastSize, 2);
        var boxX = t.X - slack;
        var boxW = t.Width + 2 * slack;

        BeginAnchor(x, ctx, boxX, plan.BoxTop, boxW, boxHeight, "Text");
        BeginWsp(x, textBox: true);
        x.WriteStartElement("wps", "spPr", WPS);
        WriteXfrm(x, boxW, boxHeight);
        WritePrstGeom(x, "rect");
        x.WriteElementString("a", "noFill", A, null);
        WriteLn(x, null);
        x.WriteEndElement(); // spPr

        x.WriteStartElement("wps", "txbx", WPS);
        x.WriteStartElement("w", "txbxContent", W);
        if (t.Lines.Count == 0)
        {
            x.WriteElementString("w", "p", W, null);
        }
        for (var k = 0; k < t.Lines.Count; k++)
            WriteLineParagraph(x, t.Lines[k], plan.LineHeights[k], slack, ctx);
        x.WriteEndElement(); // txbxContent
        x.WriteEndElement(); // txbx

        WriteEmptyBodyPr(x);
        x.WriteEndElement(); // wsp
        EndAnchor(x);
    }

    private static void WriteLineParagraph(XmlWriter x, LayoutLine line, double lineHeight, double slack, WriteContext ctx)
    {
        var slackTw = Twips(slack);
        var (left, right, jc) = line.Align switch
        {
            LayoutAlign.Left => (slackTw, 0L, "left"),
            LayoutAlign.Right => (0L, slackTw, "right"),
            LayoutAlign.Center => (0L, 0L, "center"),
            LayoutAlign.Justify => (slackTw, slackTw, "distribute"),
            _ => (slackTw, 0L, "left"),
        };

        x.WriteStartElement("w", "p", W);
        x.WriteStartElement("w", "pPr", W);
        WriteSpacing(x, Twips(lineHeight));
        x.WriteStartElement("w", "ind", W);
        x.WriteAttributeString("w", "left", W, left.ToString(CultureInfo.InvariantCulture));
        x.WriteAttributeString("w", "right", W, right.ToString(CultureInfo.InvariantCulture));
        x.WriteAttributeString("w", "firstLine", W, "0");
        x.WriteEndElement();
        WriteVal(x, "jc", jc);
        var markSize = HalfPoints(line.MaxSize > 0 ? line.MaxSize : 10).ToString(CultureInfo.InvariantCulture);
        x.WriteStartElement("w", "rPr", W);
        WriteVal(x, "sz", markSize);
        WriteVal(x, "szCs", markSize);
        x.WriteEndElement();
        x.WriteEndElement(); // pPr

        // Group consecutive runs sharing a link into one w:hyperlink.
        var i = 0;
        while (i < line.Runs.Count)
        {
            var link = line.Runs[i].Link;
            if (string.IsNullOrEmpty(link))
            {
                WriteRun(x, line.Runs[i]);
                i++;
                continue;
            }
            x.WriteStartElement("w", "hyperlink", W);
            x.WriteAttributeString("r", "id", R, LinkId(ctx, link));
            x.WriteAttributeString("w", "history", W, "1");
            while (i < line.Runs.Count && line.Runs[i].Link == link)
                WriteRun(x, line.Runs[i++]);
            x.WriteEndElement();
        }

        x.WriteEndElement(); // p
    }

    private static string LinkId(WriteContext ctx, string url)
    {
        if (ctx.LinkIds.TryGetValue(url, out var id))
            return id;
        id = $"rIdLink{ctx.LinkIds.Count + 1}";
        ctx.Main.AddHyperlinkRelationship(new Uri(url, UriKind.RelativeOrAbsolute), true, id);
        ctx.LinkIds[url] = id;
        return id;
    }

    private static void WriteRun(XmlWriter x, LayoutRun run)
    {
        x.WriteStartElement("w", "r", W);
        x.WriteStartElement("w", "rPr", W);
        if (!string.IsNullOrEmpty(run.Font))
        {
            x.WriteStartElement("w", "rFonts", W);
            x.WriteAttributeString("w", "ascii", W, run.Font);
            x.WriteAttributeString("w", "hAnsi", W, run.Font);
            x.WriteAttributeString("w", "eastAsia", W, run.Font);
            x.WriteAttributeString("w", "cs", W, run.Font);
            x.WriteEndElement();
        }
        if (run.Bold) { x.WriteElementString("w", "b", W, null); x.WriteElementString("w", "bCs", W, null); }
        if (run.Italic) { x.WriteElementString("w", "i", W, null); x.WriteElementString("w", "iCs", W, null); }
        if (run.Color is not null) WriteVal(x, "color", run.Color);
        WriteVal(x, "kern", "2"); // pair kerning at every size, like Typst
        var sz = HalfPoints(run.Size).ToString(CultureInfo.InvariantCulture);
        WriteVal(x, "sz", sz);
        WriteVal(x, "szCs", sz);
        if (run.Underline) WriteVal(x, "u", "single");
        x.WriteStartElement("w14", "ligatures", W14);
        x.WriteAttributeString("w14", "val", W14, "standard");
        x.WriteEndElement();
        x.WriteEndElement(); // rPr

        // Tabs and line breaks inside a layout run are not expected; map them defensively.
        var text = run.Text.Replace("\r", "");
        var parts = text.Split('\t');
        for (var p = 0; p < parts.Length; p++)
        {
            if (p > 0) x.WriteElementString("w", "tab", W, null);
            var segs = parts[p].Split('\n');
            for (var s = 0; s < segs.Length; s++)
            {
                if (s > 0) x.WriteElementString("w", "br", W, null);
                if (segs[s].Length == 0) continue;
                x.WriteStartElement("w", "t", W);
                x.WriteAttributeString("xml", "space", null, "preserve");
                x.WriteString(segs[s]);
                x.WriteEndElement();
            }
        }
        x.WriteEndElement(); // r
    }

    // ------------------------------------------------------------------ helpers

    private static void WriteSpacing(XmlWriter x, long lineTwips)
    {
        x.WriteStartElement("w", "spacing", W);
        x.WriteAttributeString("w", "before", W, "0");
        x.WriteAttributeString("w", "after", W, "0");
        x.WriteAttributeString("w", "line", W, Math.Max(1, lineTwips).ToString(CultureInfo.InvariantCulture));
        x.WriteAttributeString("w", "lineRule", W, "exact");
        x.WriteEndElement();
    }

    private static void WriteVal(XmlWriter x, string name, string val)
    {
        x.WriteStartElement("w", name, W);
        x.WriteAttributeString("w", "val", W, val);
        x.WriteEndElement();
    }

    public static long Emu(double pt) => (long)Math.Round(pt * EmuPerPoint, MidpointRounding.AwayFromZero);
    public static long Twips(double pt) => (long)Math.Round(pt * 20, MidpointRounding.AwayFromZero);
    public static int HalfPoints(double pt) => Math.Max(1, (int)Math.Round(pt * 2, MidpointRounding.AwayFromZero));
    private static double RoundTwips(double pt) => Math.Round(pt * 20, MidpointRounding.AwayFromZero) / 20.0;

    private static void Feed(OpenXmlPart part, string xml) => Feed(part, new UTF8Encoding(false).GetBytes(xml));

    private static void Feed(OpenXmlPart part, byte[] xml)
    {
        using var ms = new MemoryStream(xml);
        part.FeedData(ms);
    }

    /// <summary>Re-zips the package with fixed timestamps so identical input gives identical bytes.</summary>
    private static void NormalizeZip(Stream source, Stream destination)
    {
        var epoch = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
        using var src = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        using var dst = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);
        foreach (var entry in src.Entries)
        {
            var copy = dst.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            copy.LastWriteTime = epoch;
            using var from = entry.Open();
            using var to = copy.Open();
            from.CopyTo(to);
        }
    }

    // ------------------------------------------------------------------ static parts

    private const string StylesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:docDefaults>
        <w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>
        <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
        </w:docDefaults>
        <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
        <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
        <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
        </w:styles>
        """;

    private const string SettingsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:zoom w:percent="100"/>
        <w:defaultTabStop w:val="720"/>
        <w:characterSpacingControl w:val="doNotCompress"/>
        <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
        </w:settings>
        """;
}
