using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxMcp.Layout;
using Xunit;

namespace DocxMcp.Tests;

public class TechnicalDumpTests
{
    /// <summary>A heading (style with bold and size), a bulleted item, a shaded table, a header.</summary>
    private static byte[] Sample()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))),
                new Style(
                    new StyleName { Val = "heading 1" },
                    new StyleParagraphProperties(new OutlineLevel { Val = 0 }),
                    new StyleRunProperties(new Bold(), new FontSize { Val = "32" }, new Color { Val = "1F3864" }))
                { Type = StyleValues.Paragraph, StyleId = "Heading1" });

            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Header(new Paragraph(new Run(new Text("Confidential"))));

            main.Document = new Document(new Body(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }, new Justification { Val = JustificationValues.Center }),
                    new Run(new Text("Jane Doe"))),
                new Paragraph(
                    new ParagraphProperties(new NumberingProperties(new NumberingLevelReference { Val = 1 }, new NumberingId { Val = 3 })),
                    new Run(new RunProperties(new Italic()), new Text("Rust")),
                    new Run(new Text(", Typst") { Space = SpaceProcessingModeValues.Preserve })),
                new Table(new TableRow(new TableCell(
                    new TableCellProperties(new Shading { Fill = "D9E2F3", Val = ShadingPatternValues.Clear }),
                    new Paragraph(new Run(new Text("2020 – 2024"))))))));
        }
        return ms.ToArray();
    }

    private static JsonElement Dump() => JsonDocument.Parse(TechnicalDump.Dump(Sample())).RootElement;

    [Fact]
    public void Dump_ResolvesStyleChainOverDefaults()
    {
        var heading = Dump().GetProperty("body")[0];
        Assert.Equal("paragraph", heading.GetProperty("type").GetString());
        Assert.Equal("Heading1", heading.GetProperty("style").GetString());
        Assert.Equal("heading 1", heading.GetProperty("styleName").GetString());
        Assert.Equal(0, heading.GetProperty("outlineLevel").GetInt32());
        Assert.Equal("center", heading.GetProperty("align").GetString());
        var run = heading.GetProperty("runs")[0];
        Assert.Equal("Jane Doe", run.GetProperty("text").GetString());
        Assert.Equal("Calibri", run.GetProperty("font").GetString());
        Assert.Equal(16, run.GetProperty("size").GetDouble());
        Assert.True(run.GetProperty("bold").GetBoolean());
        Assert.Equal("#1f3864", run.GetProperty("color").GetString());
    }

    [Fact]
    public void Dump_KeepsListsRunsAndTables()
    {
        var root = Dump();
        var item = root.GetProperty("body")[1];
        Assert.Equal(3, item.GetProperty("list").GetProperty("id").GetInt32());
        Assert.Equal(1, item.GetProperty("list").GetProperty("level").GetInt32());
        var runs = item.GetProperty("runs");
        Assert.Equal("Rust", runs[0].GetProperty("text").GetString());
        Assert.True(runs[0].GetProperty("italic").GetBoolean());
        Assert.Equal(", Typst", runs[1].GetProperty("text").GetString());
        Assert.Equal(11, runs[1].GetProperty("size").GetDouble());

        var table = root.GetProperty("body")[2];
        Assert.Equal("table", table.GetProperty("type").GetString());
        var cell = table.GetProperty("rows")[0][0];
        Assert.Equal("#d9e2f3", cell.GetProperty("shading").GetString());
        Assert.Equal("2020 – 2024", cell.GetProperty("blocks")[0].GetProperty("runs")[0].GetProperty("text").GetString());
    }

    [Fact]
    public void Dump_ReadsHeaders()
    {
        var header = Dump().GetProperty("headers")[0];
        Assert.Equal("Confidential", header.GetProperty("runs")[0].GetProperty("text").GetString());
    }

    [Fact]
    public void Dump_IsDeterministic()
    {
        var bytes = Sample();
        Assert.Equal(TechnicalDump.Dump(bytes), TechnicalDump.Dump(bytes));
    }
}
