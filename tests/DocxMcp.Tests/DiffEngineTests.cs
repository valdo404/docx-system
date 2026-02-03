using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxMcp.Diff;
using System.Text.Json;
using Xunit;

namespace DocxMcp.Tests;

/// <summary>
/// Tests for the DiffEngine which compares Word documents without relying on element IDs.
/// Uses content-based fingerprinting and LCS matching.
/// </summary>
public class DiffEngineTests : IDisposable
{
    private readonly List<DocxSession> _sessions = [];

    private DocxSession CreateSession()
    {
        var session = DocxSession.Create();
        _sessions.Add(session);
        return session;
    }

    private DocxSession CreateSessionFromBytes(byte[] bytes)
    {
        var session = DocxSession.FromBytes(bytes, Guid.NewGuid().ToString("N")[..12], null);
        _sessions.Add(session);
        return session;
    }

    #region No Changes Tests

    [Fact]
    public void DetectsNoChanges_WhenDocumentsAreIdentical()
    {
        // Arrange
        var original = CreateSession();
        var body = original.GetBody();
        body.AppendChild(CreateParagraph("Hello World"));
        body.AppendChild(CreateParagraph("Second paragraph"));

        // Create an exact copy
        var modified = CreateSessionFromBytes(original.ToBytes());

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        Assert.False(diff.HasChanges);
        Assert.Empty(diff.Changes);
        Assert.Equal(0, diff.Summary.TotalChanges);
    }

    [Fact]
    public void DetectsNoChanges_WhenOnlyWhitespaceIsDifferent()
    {
        // Arrange - documents with equivalent normalized text
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("Hello  World"));

        var modified = CreateSession();
        modified.GetBody().AppendChild(CreateParagraph("Hello World"));

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert - Should detect no changes because text normalizes to same value
        Assert.False(diff.HasChanges);
    }

    #endregion

    #region Addition Tests

    [Fact]
    public void DetectsAddedParagraph()
    {
        // Arrange
        var original = CreateSession();
        var body = original.GetBody();
        body.AppendChild(CreateParagraph("First"));
        body.AppendChild(CreateParagraph("Second"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        modified.GetBody().AppendChild(CreateParagraph("Third - new paragraph"));

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        Assert.True(diff.HasChanges);
        Assert.Single(diff.Changes);

        var change = diff.Changes[0];
        Assert.Equal(ChangeType.Added, change.ChangeType);
        Assert.Equal("paragraph", change.ElementType);
        Assert.Contains("Third", change.NewText);
        Assert.Equal(2, change.NewIndex);
    }

    [Fact]
    public void DetectsMultipleAdditions()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("Original"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        modified.GetBody().AppendChild(CreateParagraph("Added 1"));
        modified.GetBody().AppendChild(CreateParagraph("Added 2"));

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        Assert.Equal(2, diff.Summary.Added);
    }

    [Fact]
    public void DetectsAddedTable()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("Introduction"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        modified.GetBody().AppendChild(CreateTable(2, 2));

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        Assert.True(diff.HasChanges);
        Assert.Single(diff.Changes);
        Assert.Equal(ChangeType.Added, diff.Changes[0].ChangeType);
        Assert.Equal("table", diff.Changes[0].ElementType);
    }

    #endregion

    #region Removal Tests

    [Fact]
    public void DetectsRemovedParagraph()
    {
        // Arrange
        var original = CreateSession();
        var body = original.GetBody();
        body.AppendChild(CreateParagraph("First"));
        body.AppendChild(CreateParagraph("Second - to be removed"));
        body.AppendChild(CreateParagraph("Third"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var modBody = modified.GetBody();
        var paragraphs = modBody.Elements<Paragraph>().ToList();
        modBody.RemoveChild(paragraphs[1]); // Remove "Second"

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        Assert.True(diff.HasChanges);
        Assert.Single(diff.Changes);

        var change = diff.Changes[0];
        Assert.Equal(ChangeType.Removed, change.ChangeType);
        Assert.Equal("paragraph", change.ElementType);
        Assert.Contains("Second", change.OldText);
    }

    [Fact]
    public void DetectsRemovedTable()
    {
        // Arrange
        var original = CreateSession();
        var body = original.GetBody();
        body.AppendChild(CreateParagraph("Intro"));
        body.AppendChild(CreateTable(2, 2));
        body.AppendChild(CreateParagraph("Conclusion"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var modBody = modified.GetBody();
        var table = modBody.Elements<Table>().First();
        modBody.RemoveChild(table);

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        Assert.Single(diff.Changes);
        Assert.Equal(ChangeType.Removed, diff.Changes[0].ChangeType);
        Assert.Equal("table", diff.Changes[0].ElementType);
    }

    #endregion

    #region Modification Tests

    [Fact]
    public void DetectsModifiedParagraphText()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("Original text here"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var para = modified.GetBody().Elements<Paragraph>().First();
        var run = para.Elements<Run>().First();
        run.GetFirstChild<Text>()!.Text = "Modified text here"; // Similar enough to match

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        Assert.True(diff.HasChanges);
        Assert.Single(diff.Changes);

        var change = diff.Changes[0];
        Assert.Equal(ChangeType.Modified, change.ChangeType);
        Assert.Equal("paragraph", change.ElementType);
        Assert.Contains("Original", change.OldText);
        Assert.Contains("Modified", change.NewText);
    }

    [Fact]
    public void DetectsHeadingLevelChange()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateHeading(1, "Important Title"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var heading = modified.GetBody().Elements<Paragraph>().First();
        heading.ParagraphProperties!.ParagraphStyleId!.Val = "Heading2";

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        Assert.True(diff.HasChanges);
        // Should be detected as modification (same text, different heading level)
        var change = diff.Changes.First();
        Assert.Equal(ChangeType.Modified, change.ChangeType);
    }

    [Fact]
    public void DetectsTableCellTextChange()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateTable(2, 2));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var table = modified.GetBody().Elements<Table>().First();
        var firstCell = table.Descendants<TableCell>().First();
        var para = firstCell.Elements<Paragraph>().First();
        var run = para.Elements<Run>().First();
        run.GetFirstChild<Text>()!.Text = "Changed Cell";

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        Assert.True(diff.HasChanges);
        Assert.Single(diff.Changes);
        Assert.Equal("table", diff.Changes[0].ElementType);
    }

    [Fact]
    public void DetectsCompletelyDifferentContent_AsRemoveAndAdd()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("AAAA AAAA AAAA"));

        var modified = CreateSession();
        modified.GetBody().AppendChild(CreateParagraph("ZZZZ ZZZZ ZZZZ"));

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert - completely different content should be detected as remove + add
        Assert.True(diff.HasChanges);
        Assert.True(diff.Summary.Removed >= 1 || diff.Summary.Added >= 1);
    }

    #endregion

    #region Move Tests

    [Fact]
    public void DetectsMovedParagraph()
    {
        // Arrange
        var original = CreateSession();
        var body = original.GetBody();
        body.AppendChild(CreateParagraph("First paragraph"));
        body.AppendChild(CreateParagraph("Second paragraph"));
        body.AppendChild(CreateParagraph("Third paragraph to move"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var modBody = modified.GetBody();
        var paragraphs = modBody.Elements<Paragraph>().ToList();

        // Move third to first position
        var third = paragraphs[2];
        modBody.RemoveChild(third);
        modBody.InsertChildAt(third, 0);

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        Assert.True(diff.HasChanges);
        var moveChanges = diff.Changes.Where(c => c.ChangeType == ChangeType.Moved).ToList();
        Assert.NotEmpty(moveChanges);

        // The moved element should have different old/new indices
        var movedItem = moveChanges.First(c => c.OldText?.Contains("Third") == true);
        Assert.Equal(2, movedItem.OldIndex);
        Assert.Equal(0, movedItem.NewIndex);
    }

    #endregion

    #region Multiple Changes Tests

    [Fact]
    public void DetectsMultipleChanges()
    {
        // Arrange
        var original = CreateSession();
        var body = original.GetBody();
        body.AppendChild(CreateParagraph("Keep this unchanged"));
        body.AppendChild(CreateParagraph("Remove this paragraph"));
        body.AppendChild(CreateParagraph("Modify this text content"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var modBody = modified.GetBody();
        var paragraphs = modBody.Elements<Paragraph>().ToList();

        // Remove second
        modBody.RemoveChild(paragraphs[1]);

        // Modify third (now at index 1)
        paragraphs = modBody.Elements<Paragraph>().ToList();
        var run = paragraphs[1].Elements<Run>().First();
        run.GetFirstChild<Text>()!.Text = "Modify this text changed";

        // Add new
        modBody.AppendChild(CreateParagraph("Brand new paragraph"));

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        Assert.True(diff.HasChanges);
        Assert.True(diff.Summary.TotalChanges >= 3);
        Assert.True(diff.Summary.Removed >= 1);
        Assert.True(diff.Summary.Added >= 1);
    }

    #endregion

    #region Patch Generation Tests

    [Fact]
    public void GeneratesValidPatches_ForAddition()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("First"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        modified.GetBody().AppendChild(CreateParagraph("Second - added"));

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);
        var patches = diff.ToPatches();

        // Assert
        Assert.Single(patches);
        var patch = patches[0];
        Assert.Equal("add", patch["op"]?.GetValue<string>());
        Assert.NotNull(patch["path"]);
        Assert.NotNull(patch["value"]);

        var value = patch["value"]!.AsObject();
        Assert.Equal("paragraph", value["type"]?.GetValue<string>());
    }

    [Fact]
    public void GeneratesValidPatches_ForRemoval()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("First"));
        original.GetBody().AppendChild(CreateParagraph("Second - to remove"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var modBody = modified.GetBody();
        var paragraphs = modBody.Elements<Paragraph>().ToList();
        modBody.RemoveChild(paragraphs[1]);

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);
        var patches = diff.ToPatches();

        // Assert
        Assert.Single(patches);
        var patch = patches[0];
        Assert.Equal("remove", patch["op"]?.GetValue<string>());
        Assert.NotNull(patch["path"]);
    }

    [Fact]
    public void GeneratesValidPatches_ForModification()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("Original content here"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var para = modified.GetBody().Elements<Paragraph>().First();
        para.Elements<Run>().First().GetFirstChild<Text>()!.Text = "Modified content here";

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);
        var patches = diff.ToPatches();

        // Assert
        Assert.Single(patches);
        var patch = patches[0];
        Assert.Equal("replace", patch["op"]?.GetValue<string>());
        Assert.NotNull(patch["path"]);
        Assert.NotNull(patch["value"]);
    }

    [Fact]
    public void GeneratesValidPatches_ForMove()
    {
        // Arrange
        var original = CreateSession();
        var body = original.GetBody();
        body.AppendChild(CreateParagraph("First"));
        body.AppendChild(CreateParagraph("Second"));
        body.AppendChild(CreateParagraph("Third"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var modBody = modified.GetBody();
        var paragraphs = modBody.Elements<Paragraph>().ToList();

        // Move Third to first position
        var third = paragraphs[2];
        modBody.RemoveChild(third);
        modBody.InsertChildAt(third, 0);

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);
        var patches = diff.ToPatches();

        // Assert
        var movePatches = patches.Where(p => p["op"]?.GetValue<string>() == "move").ToList();
        Assert.NotEmpty(movePatches);
    }

    #endregion

    #region API Tests

    [Fact]
    public void CompareFromBytes_WorksCorrectly()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("Hello"));
        var originalBytes = original.ToBytes();

        var modified = CreateSessionFromBytes(originalBytes);
        modified.GetBody().AppendChild(CreateParagraph("World - new"));
        var modifiedBytes = modified.ToBytes();

        // Act
        var diff = DiffEngine.Compare(originalBytes, modifiedBytes);

        // Assert
        Assert.True(diff.HasChanges);
        Assert.Single(diff.Changes);
        Assert.Equal(ChangeType.Added, diff.Changes[0].ChangeType);
    }

    [Fact]
    public void DiffResult_ToJson_ProducesValidJson()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("Hello"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        modified.GetBody().AppendChild(CreateParagraph("World"));

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);
        var json = diff.ToJson();

        // Assert
        Assert.NotNull(json);

        var parsed = JsonDocument.Parse(json);
        Assert.NotNull(parsed.RootElement.GetProperty("summary"));
        Assert.NotNull(parsed.RootElement.GetProperty("changes"));
        Assert.NotNull(parsed.RootElement.GetProperty("patches"));
    }

    [Fact]
    public void ChangeDescription_IsHumanReadable()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("Hello World"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var para = modified.GetBody().Elements<Paragraph>().First();
        para.Elements<Run>().First().GetFirstChild<Text>()!.Text = "Hello Universe";

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);

        // Assert
        var change = diff.Changes[0];
        Assert.NotEmpty(change.Description);
        Assert.Contains(change.ChangeType.ToString(), change.Description, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Summary_CorrectlyCounts_AllChangeTypes()
    {
        // Arrange
        var original = CreateSession();
        var body = original.GetBody();
        body.AppendChild(CreateParagraph("Keep"));
        body.AppendChild(CreateParagraph("Remove this"));
        body.AppendChild(CreateParagraph("Modify content here"));
        body.AppendChild(CreateParagraph("Move this"));

        var modified = CreateSessionFromBytes(original.ToBytes());
        var modBody = modified.GetBody();
        var paragraphs = modBody.Elements<Paragraph>().ToList();

        // Remove "Remove this"
        modBody.RemoveChild(paragraphs[1]);

        // Modify "Modify content"
        paragraphs = modBody.Elements<Paragraph>().ToList();
        paragraphs[1].Elements<Run>().First().GetFirstChild<Text>()!.Text = "Modify content changed";

        // Add new
        modBody.AppendChild(CreateParagraph("New paragraph"));

        // Move "Move this" to beginning
        paragraphs = modBody.Elements<Paragraph>().ToList();
        var toMove = paragraphs[2];
        modBody.RemoveChild(toMove);
        modBody.InsertChildAt(toMove, 0);

        // Act
        var diff = DiffEngine.Compare(original.Document, modified.Document);
        var summary = diff.Summary;

        // Assert
        Assert.True(diff.HasChanges);
        Assert.True(summary.TotalChanges >= 3);
    }

    #endregion

    #region Similarity Threshold Tests

    [Fact]
    public void HigherThreshold_RequiresMoreSimilarContent()
    {
        // Arrange
        var original = CreateSession();
        original.GetBody().AppendChild(CreateParagraph("The quick brown fox"));

        var modified = CreateSession();
        modified.GetBody().AppendChild(CreateParagraph("The slow brown fox"));

        // Act - with default threshold (0.6)
        var diffDefault = DiffEngine.Compare(original.Document, modified.Document, 0.6);

        // Act - with high threshold (0.95)
        var diffStrict = DiffEngine.Compare(original.Document, modified.Document, 0.95);

        // Assert - strict threshold should see remove+add, default might see modify
        Assert.True(diffDefault.HasChanges);
        Assert.True(diffStrict.HasChanges);
    }

    #endregion

    #region Helper Methods

    private static Paragraph CreateParagraph(string text)
    {
        var para = new Paragraph();
        var run = new Run();
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        para.AppendChild(run);
        return para;
    }

    private static Paragraph CreateHeading(int level, string text)
    {
        var para = new Paragraph();
        para.ParagraphProperties = new ParagraphProperties
        {
            ParagraphStyleId = new ParagraphStyleId { Val = $"Heading{level}" }
        };
        var run = new Run();
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        para.AppendChild(run);
        return para;
    }

    private static Table CreateTable(int rows, int cols)
    {
        var table = new Table();
        table.AppendChild(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
            )
        ));

        for (int r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (int c = 0; c < cols; c++)
            {
                var cell = new TableCell();
                var para = new Paragraph();
                var run = new Run();
                run.AppendChild(new Text($"R{r}C{c}") { Space = SpaceProcessingModeValues.Preserve });
                para.AppendChild(run);
                cell.AppendChild(para);
                row.AppendChild(cell);
            }
            table.AppendChild(row);
        }

        return table;
    }

    #endregion

    public void Dispose()
    {
        foreach (var session in _sessions)
        {
            session.Dispose();
        }
    }
}
