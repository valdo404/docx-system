using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxMcp.Layout;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxMcp.Tests;

public class LayoutDocxWriterTests
{
    private static string SamplePath => Path.Combine(AppContext.BaseDirectory, "TestData", "layout-sample.json");

    private static LayoutDocument LoadSample() => LayoutParser.Parse(File.ReadAllText(SamplePath));

    private static byte[] WriteSample(LayoutDocxOptions? options = null)
    {
        using var ms = new MemoryStream();
        LayoutDocxWriter.Write(LoadSample(), ms, options ?? new LayoutDocxOptions());
        return ms.ToArray();
    }

    private static WordprocessingDocument Open(byte[] bytes) =>
        WordprocessingDocument.Open(new MemoryStream(bytes), false);

    /// <summary>Top-level body paragraphs = one anchor paragraph per layout page.</summary>
    private static List<Paragraph> PageParagraphs(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToList();

    private static List<DW.Anchor> Anchors(Paragraph p) => p.Descendants<DW.Anchor>().ToList();

    private static long PosH(DW.Anchor a) => long.Parse(a.HorizontalPosition!.PositionOffset!.Text);
    private static long PosV(DW.Anchor a) => long.Parse(a.VerticalPosition!.PositionOffset!.Text);

    [Fact]
    public void Parse_Sample_ReadsAllItemTypes()
    {
        var layout = LoadSample();
        Assert.Equal(2, layout.Pages.Count);
        var items = layout.Pages[0].Items;
        Assert.Equal(6, items.Count);
        Assert.IsType<LayoutRect>(items[0]);
        Assert.IsType<LayoutImage>(items[1]);
        Assert.IsType<LayoutText>(items[2]);
        Assert.IsType<LayoutLineShape>(items[4]);
        var text = (LayoutText)items[3];
        Assert.Equal(LayoutAlign.Justify, text.Lines[0].Align);
        Assert.Equal("https://example.com/", text.Lines[1].Runs[1].Link);
        Assert.Equal("515151", text.Lines[0].Runs[0].Color);
        Assert.Null(((LayoutRect)items[0]).Stroke);
        Assert.Equal(0.8, ((LayoutRect)items[5]).Stroke!.Width);
    }

    [Fact]
    public void Parse_UnknownItemType_Throws()
    {
        var json = """{"version":1,"pages":[{"width":100,"height":100,"items":[{"type":"blob"}]}]}""";
        var ex = Assert.Throws<FormatException>(() => LayoutParser.Parse(json));
        Assert.Contains("blob", ex.Message);
    }

    [Fact]
    public void Parse_UnsupportedVersion_Throws()
    {
        Assert.Throws<FormatException>(() => LayoutParser.Parse("""{"version":2,"pages":[]}"""));
    }

    [Fact]
    public void Write_OneSectionPerPage_WithExactSizeAndZeroMargins()
    {
        using var doc = Open(WriteSample());
        var body = doc.MainDocumentPart!.Document.Body!;
        var paragraphs = PageParagraphs(doc);
        Assert.Equal(2, paragraphs.Count);

        // Page 1's section ends in its anchor paragraph, the last one is the body sectPr.
        var sections = body.Descendants<SectionProperties>().ToList();
        Assert.Equal(2, sections.Count);
        Assert.NotNull(paragraphs[0].ParagraphProperties!.SectionProperties);
        Assert.Same(body.LastChild, sections[1]);

        foreach (var s in sections)
        {
            var size = s.GetFirstChild<PageSize>()!;
            Assert.Equal(11906u, size.Width!.Value);   // 595.28pt
            Assert.Equal(16838u, size.Height!.Value);  // 841.89pt
            var mar = s.GetFirstChild<PageMargin>()!;
            Assert.Equal(0, mar.Top!.Value);
            Assert.Equal(0, mar.Bottom!.Value);
            Assert.Equal(0u, mar.Left!.Value);
            Assert.Equal(0u, mar.Right!.Value);
        }

        // Anchor paragraph is tiny: exact 1pt line, no spacing.
        var spacing = paragraphs[0].ParagraphProperties!.SpacingBetweenLines!;
        Assert.Equal("20", spacing.Line!.Value);
        Assert.Equal(LineSpacingRuleValues.Exact, spacing.LineRule!.Value);
        Assert.Equal("0", spacing.Before!.Value);
        Assert.Equal("0", spacing.After!.Value);
    }

    [Fact]
    public void Write_AnchorsPerPage_InPaintOrder()
    {
        using var doc = Open(WriteSample());
        var paragraphs = PageParagraphs(doc);
        var page1 = Anchors(paragraphs[0]);
        var page2 = Anchors(paragraphs[1]);
        Assert.Equal(6, page1.Count);
        Assert.Equal(2, page2.Count);

        var all = page1.Concat(page2).ToList();
        foreach (var a in all)
        {
            Assert.False(a.BehindDoc!.Value);
            Assert.True(a.AllowOverlap!.Value);
            Assert.NotNull(a.GetFirstChild<DW.WrapNone>());
            Assert.Equal(DW.HorizontalRelativePositionValues.Page, a.HorizontalPosition!.RelativeFrom!.Value);
            Assert.Equal(DW.VerticalRelativePositionValues.Page, a.VerticalPosition!.RelativeFrom!.Value);
        }
        // z-order follows paint order
        var z = all.Select(a => a.RelativeHeight!.Value).ToList();
        Assert.Equal(z.OrderBy(v => v).ToList(), z);
        // docPr ids unique
        var ids = all.Select(a => a.GetFirstChild<DW.DocProperties>()!.Id!.Value).ToList();
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Fact]
    public void Write_RectImageAndTextBox_AtExactEmuPositions()
    {
        var layout = LoadSample();
        using var doc = Open(WriteSample());
        var anchors = Anchors(PageParagraphs(doc)[0]);

        // rect at (0,0) 595.28 x 90
        Assert.Equal(0, PosH(anchors[0]));
        Assert.Equal(0, PosV(anchors[0]));
        Assert.Equal(595.28 * 12700, anchors[0].Extent!.Cx!.Value, 0);
        Assert.Equal(90L * 12700, anchors[0].Extent!.Cy!.Value);

        // image at (420,30) 102 x 51
        Assert.Equal(420L * 12700, PosH(anchors[1]));
        Assert.Equal(30L * 12700, PosV(anchors[1]));
        Assert.Equal(102L * 12700, anchors[1].Extent!.Cx!.Value);
        Assert.Equal(51L * 12700, anchors[1].Extent!.Cy!.Value);
        Assert.NotEmpty(anchors[1].Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>());

        // text box: widened by the horizontal slack, top from the baseline model
        var text = (LayoutText)layout.Pages[0].Items[2];
        var plan = LayoutDocxWriter.PlanText(text, BaselineModel.Default);
        Assert.Equal(LayoutDocxWriter.Emu(70.9 - 24), PosH(anchors[2]));
        Assert.Equal(LayoutDocxWriter.Emu(plan.BoxTop), PosV(anchors[2]));
        Assert.Equal(LayoutDocxWriter.Emu(300 + 48), anchors[2].Extent!.Cx!.Value);
        // default model: first line 28pt exact → baseline at 0.8*28 = 22.4pt → top = 30 + 22 - 22.4
        Assert.Equal(29.6, plan.BoxTop, 6);

        // line: horizontal, zero height
        Assert.Equal(LayoutDocxWriter.Emu(70.9), PosH(anchors[4]));
        Assert.Equal(150L * 12700, PosV(anchors[4]));
        Assert.Equal(0L, anchors[4].Extent!.Cy!.Value);
    }

    [Fact]
    public void Write_TextRuns_CarryFontsSizesColors()
    {
        using var doc = Open(WriteSample());
        var runs = doc.MainDocumentPart!.Document.Body!.Descendants<TextBoxContent>()
            .SelectMany(t => t.Descendants<Run>()).ToList();

        Assert.Equal(
            new[] { "Jane Doe", "Data engineer", "Hello ", "world",
                    ", this justified line is stretched to the full width.", "Contact: ", "example.com",
                    "right aligned", "Page two" },
            runs.Select(r => r.InnerText).ToArray());

        var name = runs[0].RunProperties!;
        Assert.Equal("Helvetica", name.RunFonts!.Ascii!.Value);
        Assert.Equal("Helvetica", name.RunFonts!.HighAnsi!.Value);
        Assert.Equal("Helvetica", name.RunFonts!.ComplexScript!.Value);
        Assert.Equal("48", name.FontSize!.Val!.Value);
        Assert.Equal("FFFFFF", name.Color!.Val!.Value);
        Assert.NotNull(name.Bold);
        Assert.Null(name.Italic);

        var role = runs[1].RunProperties!;
        Assert.NotNull(role.Italic);
        Assert.Equal("FCB912", role.Color!.Val!.Value);
        Assert.Equal("20", role.FontSize!.Val!.Value);

        Assert.Equal(UnderlineValues.Single, runs[6].RunProperties!.Underline!.Val!.Value);

        // spaces preserved
        var t = runs[2].GetFirstChild<Text>()!;
        Assert.Equal(SpaceProcessingModeValues.Preserve, t.Space!.Value);
    }

    [Fact]
    public void Write_LineParagraphs_HaveExactSpacingAndAlignment()
    {
        var layout = LoadSample();
        using var doc = Open(WriteSample());
        var boxes = doc.MainDocumentPart!.Document.Body!.Descendants<TextBoxContent>().ToList();
        Assert.Equal(3, boxes.Count);

        var text = (LayoutText)layout.Pages[0].Items[3];
        var plan = LayoutDocxWriter.PlanText(text, BaselineModel.Default);
        var paras = boxes[1].Elements<Paragraph>().ToList();
        Assert.Equal(3, paras.Count);
        for (var i = 0; i < paras.Count; i++)
        {
            var sp = paras[i].ParagraphProperties!.SpacingBetweenLines!;
            Assert.Equal(LineSpacingRuleValues.Exact, sp.LineRule!.Value);
            Assert.Equal(LayoutDocxWriter.Twips(plan.LineHeights[i]).ToString(), sp.Line!.Value);
            Assert.Equal("0", sp.Before!.Value);
            Assert.Equal("0", sp.After!.Value);
        }
        Assert.Equal(JustificationValues.Distribute, paras[0].ParagraphProperties!.Justification!.Val!.Value);
        Assert.Equal(JustificationValues.Left, paras[1].ParagraphProperties!.Justification!.Val!.Value);
        Assert.Equal(JustificationValues.Right, paras[2].ParagraphProperties!.Justification!.Val!.Value);
        // slack compensated by indents: left line starts at the layout x
        Assert.Equal("480", paras[1].ParagraphProperties!.Indentation!.Left!.Value);
        Assert.Equal("480", paras[2].ParagraphProperties!.Indentation!.Right!.Value);

        var center = boxes[2].Elements<Paragraph>().Single();
        Assert.Equal(JustificationValues.Center, center.ParagraphProperties!.Justification!.Val!.Value);
    }

    [Fact]
    public void Write_Hyperlink_HasExternalRelationship()
    {
        using var doc = Open(WriteSample());
        var main = doc.MainDocumentPart!;
        var link = main.Document.Body!.Descendants<Hyperlink>().Single();
        Assert.Equal("example.com", link.InnerText);
        var rel = main.HyperlinkRelationships.Single(r => r.Id == link.Id!.Value);
        Assert.True(rel.IsExternal);
        Assert.Equal("https://example.com/", rel.Uri.ToString());
    }

    [Fact]
    public void Write_ImagePart_IsEmbedded()
    {
        using var doc = Open(WriteSample());
        var main = doc.MainDocumentPart!;
        var blip = main.Document.Body!.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Single();
        var part = (ImagePart)main.GetPartById(blip.Embed!.Value!);
        Assert.Equal("image/png", part.ContentType);
    }

    [Fact]
    public void Write_Output_PassesOpenXmlValidation()
    {
        using var doc = Open(WriteSample());
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        var errors = validator.Validate(doc).ToList();
        Assert.True(errors.Count == 0,
            string.Join("\n", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    [Fact]
    public void Write_RunSpacing_EmitsCharacterSpacingInTwips()
    {
        var json = """
            {"version":1,"pages":[{"width":200,"height":100,"items":[
              {"type":"text","x":10,"y":10,"width":180,"height":12,"lines":[
                {"baseline":9.6,"height":12,"align":"left","runs":[
                  {"text":"Hello","font":"Helvetica","size":10},
                  {"text":" ","font":"Helvetica","size":10,"spacing":2.35},
                  {"text":"tight","font":"Helvetica","size":10,"spacing":-0.4},
                  {"text":"x","font":"Helvetica","size":10,"spacing":5000},
                  {"text":"y","font":"Helvetica","size":10,"spacing":0}
                ]}]}]}]}
            """;
        var layout = LayoutParser.Parse(json);
        Assert.Equal(2.35, layout.Pages[0].Items.OfType<LayoutText>().Single().Lines[0].Runs[1].Spacing);

        using var ms = new MemoryStream();
        LayoutDocxWriter.Write(layout, ms);
        using var doc = Open(ms.ToArray());
        var runs = doc.MainDocumentPart!.Document.Body!.Descendants<TextBoxContent>()
            .SelectMany(t => t.Descendants<Run>()).ToList();
        Assert.Equal(5, runs.Count);
        Assert.Null(runs[0].RunProperties!.Spacing);
        Assert.Equal(47, runs[1].RunProperties!.Spacing!.Val!.Value);     // 2.35pt = 47 twips
        Assert.Equal(-8, runs[2].RunProperties!.Spacing!.Val!.Value);     // negative allowed
        Assert.Equal(31680, runs[3].RunProperties!.Spacing!.Val!.Value);  // clamped to 1584pt
        Assert.Null(runs[4].RunProperties!.Spacing);                      // 0 → omitted
        Assert.Equal(JustificationValues.Left,
            runs[0].Ancestors<Paragraph>().First().ParagraphProperties!.Justification!.Val!.Value);

        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(doc).ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    [Fact]
    public void Write_IsDeterministic()
    {
        var a = WriteSample();
        Thread.Sleep(1100); // zip timestamps have a 2 s resolution; make sure time moved on
        var b = WriteSample();
        Assert.Equal(a, b);
    }

    [Theory]
    [InlineData(0.8)]
    [InlineData(0.75)]
    [InlineData(0.9)]
    public void PlanText_PutsEveryBaselineOnTarget(double ratio)
    {
        var model = new BaselineModel(ratio);
        var text = (LayoutText)LoadSample().Pages[0].Items[3];
        var plan = LayoutDocxWriter.PlanText(text, model);

        var top = plan.BoxTop;
        for (var k = 0; k < text.Lines.Count; k++)
        {
            var h = plan.LineHeights[k];
            var baseline = top + model.BaselineOffset(h, text.Lines[k].MaxSize);
            // exact line heights are rounded to twips (1/20 pt)
            Assert.InRange(baseline - (text.Y + text.Lines[k].Baseline), -0.05, 0.05);
            top += h;
        }
    }

    [Fact]
    public void BaselineModel_DefaultFormula_AndInverse()
    {
        var m = BaselineModel.Default;
        Assert.Equal(11.2, m.BaselineOffset(14, 10), 9);
        Assert.Equal(4.8, m.BaselineOffset(6, 10), 9);   // measured: also when the line is smaller than the font
        Assert.Equal(14, m.SolveLineHeight(11.2, 10)!.Value, 6);
        Assert.Equal(5, m.SolveLineHeight(4.0, 10)!.Value, 6);
        Assert.Null(m.SolveLineHeight(0.0, 10));         // unreachable: baseline at the line top
    }
}

public class LayoutToolsTests
{
    [Fact]
    public void DocumentFromLayout_ReturnsBase64Docx_OrWritesFile()
    {
        var json = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "TestData", "layout-sample.json"));

        var b64 = DocxMcp.Tools.LayoutTools.DocumentFromLayout(json);
        var bytes = Convert.FromBase64String(b64);
        using (var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false))
            Assert.Equal(2, doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Count());

        var path = Path.Combine(Path.GetTempPath(), $"layout-{Guid.NewGuid():N}.docx");
        try
        {
            var msg = DocxMcp.Tools.LayoutTools.DocumentFromLayout(json, path);
            Assert.Contains("2 page(s)", msg);
            Assert.Equal(bytes, File.ReadAllBytes(path));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DocumentFromLayout_InvalidJson_ThrowsMcpException()
    {
        Assert.Throws<ModelContextProtocol.McpException>(
            () => DocxMcp.Tools.LayoutTools.DocumentFromLayout("""{"version":1,"pages":[{"width":1}]}"""));
    }
}
