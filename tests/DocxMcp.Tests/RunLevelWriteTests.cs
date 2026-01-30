using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxMcp.Helpers;
using DocxMcp.Paths;
using System.Text.Json;
using Xunit;

namespace DocxMcp.Tests;

public class RunLevelWriteTests : IDisposable
{
    private readonly DocxSession _session;

    public RunLevelWriteTests()
    {
        _session = DocxSession.Create();
    }

    [Fact]
    public void CreateParagraphWithRunsArray()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "runs": [
                {"text": "Hello ", "style": {"bold": true}},
                {"text": "World"}
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var runs = p.Elements<Run>().ToList();

        Assert.Equal(2, runs.Count);
        Assert.Equal("Hello ", runs[0].InnerText);
        Assert.NotNull(runs[0].RunProperties?.Bold);
        Assert.Equal("World", runs[1].InnerText);
        Assert.Null(runs[1].RunProperties);
    }

    [Fact]
    public void CreateParagraphWithTabRun()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "runs": [
                {"text": "Title"},
                {"tab": true},
                {"text": "Company"}
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var runs = p.Elements<Run>().ToList();

        Assert.Equal(3, runs.Count);
        Assert.Equal("Title", runs[0].InnerText);
        Assert.NotNull(runs[1].GetFirstChild<TabChar>());
        Assert.Equal("Company", runs[2].InnerText);
    }

    [Fact]
    public void CreateParagraphWithBreakRun()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "runs": [
                {"text": "Before"},
                {"break": "line"},
                {"text": "After"}
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var runs = p.Elements<Run>().ToList();

        Assert.Equal(3, runs.Count);
        var brk = runs[1].GetFirstChild<Break>();
        Assert.NotNull(brk);
        Assert.Equal(BreakValues.TextWrapping, brk!.Type?.Value);
    }

    [Fact]
    public void CreateParagraphWithPageBreakRun()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "runs": [
                {"break": "page"}
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var runs = p.Elements<Run>().ToList();

        Assert.Single(runs);
        var brk = runs[0].GetFirstChild<Break>();
        Assert.NotNull(brk);
        Assert.Equal(BreakValues.Page, brk!.Type?.Value);
    }

    [Fact]
    public void CreateParagraphWithFullRunStyling()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "runs": [
                {
                    "text": "Styled",
                    "style": {
                        "bold": true,
                        "italic": true,
                        "underline": true,
                        "strike": true,
                        "font_size": 16,
                        "font_name": "Arial",
                        "color": "FF0000"
                    }
                }
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var run = p.Elements<Run>().Single();
        var rp = run.RunProperties;

        Assert.NotNull(rp);
        Assert.NotNull(rp!.Bold);
        Assert.NotNull(rp.Italic);
        Assert.NotNull(rp.Underline);
        Assert.NotNull(rp.Strike);
        Assert.Equal("32", rp.FontSize?.Val?.Value); // 16pt = 32 half-points
        Assert.Equal("Arial", rp.RunFonts?.Ascii?.Value);
        Assert.Equal("FF0000", rp.Color?.Val?.Value);
    }

    [Fact]
    public void CreateParagraphWithHighlightAndVerticalAlign()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "runs": [
                {"text": "Normal"},
                {"text": "2", "style": {"vertical_align": "superscript"}},
                {"text": " highlighted", "style": {"highlight": "yellow"}}
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var runs = p.Elements<Run>().ToList();

        Assert.Equal(3, runs.Count);
        Assert.Equal(VerticalPositionValues.Superscript,
            runs[1].RunProperties?.VerticalTextAlignment?.Val?.Value);
        Assert.Equal(HighlightColorValues.Yellow,
            runs[2].RunProperties?.Highlight?.Val?.Value);
    }

    [Fact]
    public void CreateHeadingWithRunsArray()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "heading",
            "level": 2,
            "runs": [
                {"text": "Job Title  ", "style": {"color": "2E5496"}},
                {"tab": true},
                {"text": "Company", "style": {"bold": true}},
                {"tab": true},
                {"text": "Jan 2020", "style": {"italic": true}}
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);

        // Verify it's a heading
        Assert.Equal("Heading2", p.ParagraphProperties?.ParagraphStyleId?.Val?.Value);

        // Verify runs
        var runs = p.Elements<Run>().ToList();
        Assert.Equal(5, runs.Count);

        Assert.Equal("Job Title  ", runs[0].InnerText);
        Assert.Equal("2E5496", runs[0].RunProperties?.Color?.Val?.Value);

        Assert.NotNull(runs[1].GetFirstChild<TabChar>());

        Assert.Equal("Company", runs[2].InnerText);
        Assert.NotNull(runs[2].RunProperties?.Bold);

        Assert.NotNull(runs[3].GetFirstChild<TabChar>());

        Assert.Equal("Jan 2020", runs[4].InnerText);
        Assert.NotNull(runs[4].RunProperties?.Italic);
    }

    [Fact]
    public void CreateHeadingWithFlatTextFallback()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "heading",
            "level": 1,
            "text": "Simple Title"
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        Assert.Equal("Heading1", p.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Equal("Simple Title", p.InnerText);
    }

    [Fact]
    public void CreateParagraphWithProperties()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "properties": {
                "alignment": "center",
                "spacing_before": 120,
                "spacing_after": 240
            },
            "runs": [
                {"text": "Centered paragraph"}
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var pp = p.ParagraphProperties;

        Assert.NotNull(pp);
        Assert.Equal(JustificationValues.Center, pp!.Justification?.Val?.Value);
        Assert.Equal("120", pp.SpacingBetweenLines?.Before?.Value);
        Assert.Equal("240", pp.SpacingBetweenLines?.After?.Value);
        Assert.Equal("Centered paragraph", p.InnerText);
    }

    [Fact]
    public void CreateParagraphWithIndentation()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "properties": {
                "indent_left": 720,
                "indent_right": 360,
                "indent_first_line": 360
            },
            "runs": [
                {"text": "Indented paragraph"}
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var indent = p.ParagraphProperties?.Indentation;

        Assert.NotNull(indent);
        Assert.Equal("720", indent!.Left?.Value);
        Assert.Equal("360", indent.Right?.Value);
        Assert.Equal("360", indent.FirstLine?.Value);
    }

    [Fact]
    public void CreateParagraphWithTabStops()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "properties": {
                "tabs": [
                    {"position": 4680, "alignment": "center"},
                    {"position": 9360, "alignment": "right", "leader": "dot"}
                ]
            },
            "runs": [
                {"text": "Left"},
                {"tab": true},
                {"text": "Center"},
                {"tab": true},
                {"text": "Right"}
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var tabs = p.ParagraphProperties?.Tabs?.Elements<TabStop>().ToList();

        Assert.NotNull(tabs);
        Assert.Equal(2, tabs!.Count);
        Assert.Equal(4680, tabs[0].Position?.Value);
        Assert.Equal(TabStopValues.Center, tabs[0].Val?.Value);
        Assert.Equal(9360, tabs[1].Position?.Value);
        Assert.Equal(TabStopValues.Right, tabs[1].Val?.Value);
        Assert.Equal(TabStopLeaderCharValues.Dot, tabs[1].Leader?.Value);
    }

    [Fact]
    public void CreateHeadingWithPropertiesMerge()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "heading",
            "level": 2,
            "properties": {
                "alignment": "right",
                "tabs": [
                    {"position": 9360, "alignment": "right"}
                ]
            },
            "runs": [
                {"text": "Title"},
                {"tab": true},
                {"text": "Date"}
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var pp = p.ParagraphProperties;

        Assert.NotNull(pp);
        // Heading style is preserved
        Assert.Equal("Heading2", pp!.ParagraphStyleId?.Val?.Value);
        // Additional properties are merged
        Assert.Equal(JustificationValues.Right, pp.Justification?.Val?.Value);
        Assert.NotNull(pp.Tabs);
        Assert.Single(pp.Tabs!.Elements<TabStop>());
    }

    [Fact]
    public void LegacyFlatTextStillWorks()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "text": "Simple text",
            "style": {"bold": true}
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        Assert.Equal("Simple text", p.InnerText);

        var run = p.Elements<Run>().Single();
        Assert.NotNull(run.RunProperties?.Bold);
    }

    [Fact]
    public void CreateRunDirectly()
    {
        var runJson = JsonDocument.Parse("""
            {"text": "Hello", "style": {"bold": true, "color": "FF0000"}}
        """).RootElement;

        var run = ElementFactory.CreateRun(runJson);
        Assert.Equal("Hello", run.InnerText);
        Assert.NotNull(run.RunProperties?.Bold);
        Assert.Equal("FF0000", run.RunProperties?.Color?.Val?.Value);
    }

    [Fact]
    public void CreateTabRunDirectly()
    {
        var runJson = JsonDocument.Parse("""{"tab": true}""").RootElement;

        var run = ElementFactory.CreateRun(runJson);
        Assert.NotNull(run.GetFirstChild<TabChar>());
    }

    [Fact]
    public void CreateTabRunWithStyling()
    {
        var runJson = JsonDocument.Parse("""
            {"tab": true, "style": {"font_size": 11}}
        """).RootElement;

        var run = ElementFactory.CreateRun(runJson);
        Assert.NotNull(run.GetFirstChild<TabChar>());
        Assert.Equal("22", run.RunProperties?.FontSize?.Val?.Value);
    }

    [Fact]
    public void RoundTripParagraphWithMultipleStyledRuns()
    {
        var mainPart = _session.Document.MainDocumentPart!;

        // Create a paragraph with runs matching the CV heading format
        var value = JsonDocument.Parse("""
        {
            "type": "heading",
            "level": 2,
            "runs": [
                {"text": "Lead Scala Developer  ", "style": {"color": "2E5496"}},
                {"tab": true, "style": {"color": "2E5496"}},
                {"text": "_", "style": {"bold": true}},
                {"text": "Powerspace"},
                {"text": "_ ", "style": {"bold": true}},
                {"text": "_", "style": {"bold": true}},
                {"text": "August 2015 to March 2019"},
                {"text": " ", "style": {"bold": true, "italic": true, "font_size": 11}}
            ]
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);

        Assert.Equal("Heading2", p.ParagraphProperties?.ParagraphStyleId?.Val?.Value);

        var runs = p.Elements<Run>().ToList();
        Assert.Equal(8, runs.Count);

        // First run: colored text
        Assert.Equal("Lead Scala Developer  ", runs[0].InnerText);
        Assert.Equal("2E5496", runs[0].RunProperties?.Color?.Val?.Value);

        // Second run: tab
        Assert.NotNull(runs[1].GetFirstChild<TabChar>());
        Assert.Equal("2E5496", runs[1].RunProperties?.Color?.Val?.Value);

        // Third run: bold underscore
        Assert.Equal("_", runs[2].InnerText);
        Assert.NotNull(runs[2].RunProperties?.Bold);

        // Fourth run: plain text
        Assert.Equal("Powerspace", runs[3].InnerText);
        Assert.Null(runs[3].RunProperties);

        // Last run: bold + italic + sized
        Assert.NotNull(runs[7].RunProperties?.Bold);
        Assert.NotNull(runs[7].RunProperties?.Italic);
        Assert.Equal("22", runs[7].RunProperties?.FontSize?.Val?.Value);
    }

    [Fact]
    public void ParagraphWithEmptyRunsArrayCreatesNoParagraphContent()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "runs": []
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        Assert.Empty(p.Elements<Run>());
    }

    [Fact]
    public void ParagraphPropertiesWithLineSpacing()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "properties": {
                "line_spacing": 360,
                "spacing_before": 0,
                "spacing_after": 0
            },
            "text": "Double-spaced"
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var spacing = p.ParagraphProperties?.SpacingBetweenLines;

        Assert.NotNull(spacing);
        Assert.Equal("360", spacing!.Line?.Value);
        Assert.Equal("0", spacing.Before?.Value);
        Assert.Equal("0", spacing.After?.Value);
    }

    [Fact]
    public void ParagraphPropertiesWithHangingIndent()
    {
        var mainPart = _session.Document.MainDocumentPart!;
        var value = JsonDocument.Parse("""
        {
            "type": "paragraph",
            "properties": {
                "indent_left": 720,
                "indent_hanging": 360
            },
            "text": "Hanging indent paragraph"
        }
        """).RootElement;

        var element = ElementFactory.CreateFromJson(value, mainPart);
        var p = Assert.IsType<Paragraph>(element);
        var indent = p.ParagraphProperties?.Indentation;

        Assert.NotNull(indent);
        Assert.Equal("720", indent!.Left?.Value);
        Assert.Equal("360", indent.Hanging?.Value);
    }

    public void Dispose()
    {
        _session.Dispose();
    }
}
