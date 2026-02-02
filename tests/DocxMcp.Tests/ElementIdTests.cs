using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxMcp.Helpers;
using DocxMcp.Paths;
using Xunit;

namespace DocxMcp.Tests;

public class ElementIdTests : IDisposable
{
    private readonly MemoryStream _stream;
    private readonly WordprocessingDocument _doc;

    public ElementIdTests()
    {
        _stream = new MemoryStream();
        _doc = WordprocessingDocument.Create(_stream, WordprocessingDocumentType.Document);
        var mainPart = _doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var body = mainPart.Document.Body!;
        body.AppendChild(new Paragraph(new Run(new Text("First"))));
        body.AppendChild(new Paragraph(new Run(new Text("Second"))));

        var table = new Table(
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("A")))),
                new TableCell(new Paragraph(new Run(new Text("B"))))),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("C")))),
                new TableCell(new Paragraph(new Run(new Text("D"))))));
        body.AppendChild(table);
    }

    [Fact]
    public void GenerateId_ValidRange()
    {
        var id = ElementIdManager.GenerateId();
        var value = Convert.ToInt32(id, 16);
        Assert.True(value > 0);
        Assert.True(value < 0x7FFFFFFF);
    }

    [Fact]
    public void GenerateId_8CharHex()
    {
        var id = ElementIdManager.GenerateId();
        Assert.Equal(8, id.Length);
        Assert.True(id.All(c => "0123456789ABCDEF".Contains(c)));
    }

    [Fact]
    public void GenerateId_NoCollisions()
    {
        var existing = new HashSet<string>();
        for (int i = 0; i < 1000; i++)
        {
            var id = ElementIdManager.GenerateId(existing);
            Assert.DoesNotContain(id, existing.Where(e => e != id));
        }
        Assert.Equal(1000, existing.Count);
    }

    [Fact]
    public void EnsureNamespace_DeclaresNamespaces()
    {
        ElementIdManager.EnsureNamespace(_doc);

        var document = _doc.MainDocumentPart!.Document!;
        var nsDecls = document.NamespaceDeclarations.ToDictionary(d => d.Key, d => d.Value);

        Assert.True(nsDecls.ContainsKey("dmcp"));
        Assert.Equal("http://docx-mcp.dev/id", nsDecls["dmcp"]);
        Assert.True(nsDecls.ContainsKey("mc"));
    }

    [Fact]
    public void EnsureNamespace_McIgnorableIncludesDmcp()
    {
        ElementIdManager.EnsureNamespace(_doc);

        var document = _doc.MainDocumentPart!.Document!;
        var attrs = document.GetAttributes();
        var ignorable = attrs.FirstOrDefault(a => a.LocalName == "Ignorable");

        Assert.NotNull(ignorable.Value);
        Assert.Contains("dmcp", ignorable.Value.Split(' '));
    }

    [Fact]
    public void EnsureAllIds_AssignsToAllElementTypes()
    {
        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        var body = _doc.MainDocumentPart!.Document!.Body!;

        // Check paragraphs
        foreach (var p in body.Descendants<Paragraph>())
        {
            Assert.NotNull(ElementIdManager.GetId(p));
        }

        // Check table
        foreach (var t in body.Descendants<Table>())
        {
            Assert.NotNull(ElementIdManager.GetId(t));
        }

        // Check rows
        foreach (var tr in body.Descendants<TableRow>())
        {
            Assert.NotNull(ElementIdManager.GetId(tr));
        }

        // Check cells
        foreach (var tc in body.Descendants<TableCell>())
        {
            Assert.NotNull(ElementIdManager.GetId(tc));
        }

        // Check runs
        foreach (var r in body.Descendants<Run>())
        {
            Assert.NotNull(ElementIdManager.GetId(r));
        }
    }

    [Fact]
    public void EnsureAllIds_PreservesExisting()
    {
        // Manually set an ID on the first paragraph
        var firstPara = _doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        ElementIdManager.SetDmcpId(firstPara, "DEADBEEF");

        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        // Verify the manually-set ID is preserved
        Assert.Equal("DEADBEEF", ElementIdManager.GetId(firstPara));
    }

    [Fact]
    public void EnsureAllIds_ParagraphasHaveParaIdAndTextId()
    {
        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        var body = _doc.MainDocumentPart!.Document!.Body!;
        foreach (var p in body.Elements<Paragraph>())
        {
            Assert.NotNull(p.ParagraphId);
            Assert.False(string.IsNullOrEmpty(p.ParagraphId!.Value));
            Assert.NotNull(p.TextId);
            Assert.False(string.IsNullOrEmpty(p.TextId!.Value));
        }
    }

    [Fact]
    public void EnsureAllIds_TableRowsHaveParaId()
    {
        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        var body = _doc.MainDocumentPart!.Document!.Body!;
        foreach (var tr in body.Descendants<TableRow>())
        {
            Assert.NotNull(tr.ParagraphId);
            Assert.False(string.IsNullOrEmpty(tr.ParagraphId!.Value));
            Assert.NotNull(tr.TextId);
            Assert.False(string.IsNullOrEmpty(tr.TextId!.Value));
        }
    }

    [Fact]
    public void EnsureAllIds_AllIdsUnique()
    {
        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        var body = _doc.MainDocumentPart!.Document!.Body!;
        var allIds = new HashSet<string>();

        foreach (var element in body.Descendants())
        {
            var id = ElementIdManager.GetId(element);
            if (id is not null)
            {
                Assert.True(allIds.Add(id), $"Duplicate ID: {id}");
            }
        }

        // Should have at least as many IDs as paragraphs + tables + rows + cells + runs
        Assert.True(allIds.Count > 0);
    }

    [Fact]
    public void AssignId_SingleElement()
    {
        var paragraph = new Paragraph(new Run(new Text("test")));
        ElementIdManager.AssignId(paragraph);

        var id = ElementIdManager.GetId(paragraph);
        Assert.NotNull(id);
        Assert.Equal(8, id.Length);

        // Also sets w14:paraId for paragraphs
        Assert.NotNull(paragraph.ParagraphId);
    }

    [Fact]
    public void AssignId_TableRow()
    {
        var row = new TableRow(new TableCell(new Paragraph(new Run(new Text("test")))));
        ElementIdManager.AssignId(row);

        var id = ElementIdManager.GetId(row);
        Assert.NotNull(id);

        // Also sets w14:paraId for rows
        Assert.NotNull(row.ParagraphId);
    }

    [Fact]
    public void GetId_FallsBackToParaId()
    {
        // Simulate a paragraph that only has w14:paraId (Word stripped dmcp:id)
        var paragraph = new Paragraph(new Run(new Text("test")));
        paragraph.ParagraphId = new HexBinaryValue("1A2B3C4D");

        var id = ElementIdManager.GetId(paragraph);
        Assert.Equal("1A2B3C4D", id);
    }

    [Fact]
    public void ReDerivesFromParaId_WhenDmcpIdStripped()
    {
        // Assign IDs
        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        var body = _doc.MainDocumentPart!.Document!.Body!;
        var firstPara = body.Elements<Paragraph>().First();
        var originalParaId = firstPara.ParagraphId!.Value;

        // Simulate Word stripping dmcp:id attributes
        foreach (var element in body.Descendants())
        {
            var attrs = element.GetAttributes().ToList();
            var dmcpAttr = attrs.FirstOrDefault(
                a => a.LocalName == "id" && a.NamespaceUri == ElementIdManager.DmcpNamespace);
            if (!string.IsNullOrEmpty(dmcpAttr.Value))
            {
                element.RemoveAttribute(dmcpAttr.LocalName, dmcpAttr.NamespaceUri);
            }
        }

        // Verify dmcp:id is gone from paragraphs
        Assert.Null(ElementIdManager.GetDmcpId(firstPara));

        // Re-run EnsureAllIds — should re-derive from w14:paraId
        ElementIdManager.EnsureAllIds(_doc);

        // dmcp:id should now match the original w14:paraId
        var newDmcpId = ElementIdManager.GetDmcpId(firstPara);
        Assert.Equal(originalParaId, newDmcpId);
    }

    [Fact]
    public void IdSelector_ParsesCorrectly()
    {
        var path = DocxPath.Parse("/body/paragraph[id='1A2B3C4D']");
        var seg = Assert.IsType<ParagraphSegment>(path.Segments[1]);
        var sel = Assert.IsType<IdSelector>(seg.Selector);
        Assert.Equal("1A2B3C4D", sel.Id);
    }

    [Fact]
    public void IdSelector_CaseInsensitiveParsing()
    {
        var path = DocxPath.Parse("/body/paragraph[id='abcdef01']");
        var seg = Assert.IsType<ParagraphSegment>(path.Segments[1]);
        var sel = Assert.IsType<IdSelector>(seg.Selector);
        Assert.Equal("ABCDEF01", sel.Id); // Normalized to uppercase
    }

    [Fact]
    public void IdSelector_OnTable()
    {
        var path = DocxPath.Parse("/body/table[id='12345678']");
        var seg = Assert.IsType<TableSegment>(path.Segments[1]);
        Assert.IsType<IdSelector>(seg.Selector);
    }

    [Fact]
    public void IdSelector_OnRow()
    {
        var path = DocxPath.Parse("/body/table[0]/row[id='AABBCCDD']");
        var seg = Assert.IsType<RowSegment>(path.Segments[2]);
        Assert.IsType<IdSelector>(seg.Selector);
    }

    [Fact]
    public void IdSelector_OnCell()
    {
        var path = DocxPath.Parse("/body/table[0]/row[0]/cell[id='11223344']");
        var seg = Assert.IsType<CellSegment>(path.Segments[3]);
        Assert.IsType<IdSelector>(seg.Selector);
    }

    [Fact]
    public void IdSelector_OnRun()
    {
        var path = DocxPath.Parse("/body/paragraph[0]/run[id='AABB1122']");
        var seg = Assert.IsType<RunSegment>(path.Segments[2]);
        Assert.IsType<IdSelector>(seg.Selector);
    }

    [Fact]
    public void IdSelector_ResolveParagraphById()
    {
        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        var body = _doc.MainDocumentPart!.Document!.Body!;
        var firstPara = body.Elements<Paragraph>().First();
        var id = ElementIdManager.GetId(firstPara)!;

        var path = DocxPath.Parse($"/body/paragraph[id='{id}']");
        var results = PathResolver.Resolve(path, _doc);

        Assert.Single(results);
        Assert.Same(firstPara, results[0]);
    }

    [Fact]
    public void IdSelector_ResolveTableById()
    {
        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        var body = _doc.MainDocumentPart!.Document!.Body!;
        var table = body.Elements<Table>().First();
        var id = ElementIdManager.GetId(table)!;

        var path = DocxPath.Parse($"/body/table[id='{id}']");
        var results = PathResolver.Resolve(path, _doc);

        Assert.Single(results);
        Assert.Same(table, results[0]);
    }

    [Fact]
    public void IdSelector_ResolveRowById()
    {
        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        var body = _doc.MainDocumentPart!.Document!.Body!;
        var table = body.Elements<Table>().First();
        var firstRow = table.Elements<TableRow>().First();
        var rowId = ElementIdManager.GetId(firstRow)!;
        var tableId = ElementIdManager.GetId(table)!;

        var path = DocxPath.Parse($"/body/table[id='{tableId}']/row[id='{rowId}']");
        var results = PathResolver.Resolve(path, _doc);

        Assert.Single(results);
        Assert.Same(firstRow, results[0]);
    }

    [Fact]
    public void IdSelector_ResolveCellById()
    {
        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        var body = _doc.MainDocumentPart!.Document!.Body!;
        var cell = body.Descendants<TableCell>().First();
        var id = ElementIdManager.GetId(cell)!;

        var row = cell.Parent as TableRow;
        var rowId = ElementIdManager.GetId(row!)!;
        var table = row!.Parent as Table;
        var tableId = ElementIdManager.GetId(table!)!;

        var path = DocxPath.Parse($"/body/table[id='{tableId}']/row[id='{rowId}']/cell[id='{id}']");
        var results = PathResolver.Resolve(path, _doc);

        Assert.Single(results);
        Assert.Same(cell, results[0]);
    }

    [Fact]
    public void IdSelector_ResolveRunById()
    {
        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        var body = _doc.MainDocumentPart!.Document!.Body!;
        var firstPara = body.Elements<Paragraph>().First();
        var firstRun = firstPara.Elements<Run>().First();
        var runId = ElementIdManager.GetId(firstRun)!;
        var paraId = ElementIdManager.GetId(firstPara)!;

        var path = DocxPath.Parse($"/body/paragraph[id='{paraId}']/run[id='{runId}']");
        var results = PathResolver.Resolve(path, _doc);

        Assert.Single(results);
        Assert.Same(firstRun, results[0]);
    }

    [Fact]
    public void IdSelector_NotFoundThrows()
    {
        ElementIdManager.EnsureNamespace(_doc);
        ElementIdManager.EnsureAllIds(_doc);

        var path = DocxPath.Parse("/body/paragraph[id='00000000']");
        Assert.Throws<InvalidOperationException>(() => PathResolver.Resolve(path, _doc));
    }

    [Fact]
    public void RoundTrip_CreateQueryPatch()
    {
        // Create a session, assign IDs, query by ID
        var session = DocxSession.Create();
        try
        {
            var body = session.GetBody();

            // Add a paragraph
            var para = new Paragraph(new Run(new Text("Test paragraph")));
            ElementIdManager.AssignId(para);
            body.AppendChild(para);

            var id = ElementIdManager.GetId(para)!;

            // Resolve by ID
            var path = DocxPath.Parse($"/body/paragraph[id='{id}']");
            var results = PathResolver.Resolve(path, session.Document);
            Assert.Single(results);
            Assert.Equal("Test paragraph", results[0].InnerText);
        }
        finally
        {
            session.Dispose();
        }
    }

    [Fact]
    public void DocxSession_Create_AssignsIds()
    {
        var session = DocxSession.Create();
        try
        {
            // Namespace should be set
            var document = session.Document.MainDocumentPart!.Document!;
            var nsDecls = document.NamespaceDeclarations.ToDictionary(d => d.Key, d => d.Value);
            Assert.True(nsDecls.ContainsKey("dmcp"));
        }
        finally
        {
            session.Dispose();
        }
    }

    public void Dispose()
    {
        _doc.Dispose();
        _stream.Dispose();
    }
}
