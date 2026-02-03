using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxMcp.ExternalChanges;
using DocxMcp.Persistence;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace DocxMcp.Tests;

/// <summary>
/// Tests that sync-external does not create duplicate WAL entries
/// when called multiple times without actual file changes.
/// This was the bug: ID reassignment caused hash mismatches even when
/// the external file hadn't changed.
/// </summary>
public class SyncDuplicateTests : IDisposable
{
    private readonly string _tempDir;
    private readonly string _tempFile;
    private readonly SessionStore _store;
    private readonly SessionManager _sessionManager;
    private readonly ExternalChangeTracker _tracker;

    public SyncDuplicateTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"docx-mcp-sync-dup-test-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _tempFile = Path.Combine(_tempDir, "test.docx");

        // Create test document
        CreateTestDocx(_tempFile, "Test content");

        _store = new SessionStore(NullLogger<SessionStore>.Instance, _tempDir);
        _sessionManager = new SessionManager(_store, NullLogger<SessionManager>.Instance);
        _tracker = new ExternalChangeTracker(_sessionManager, NullLogger<ExternalChangeTracker>.Instance);
    }

    [Fact]
    public void SyncExternalChanges_CalledTwice_OnlyCreatesOneWalEntry()
    {
        // Arrange
        var session = _sessionManager.Open(_tempFile);

        // Act - first sync (may or may not have changes depending on ID assignment)
        var result1 = _tracker.SyncExternalChanges(session.Id);

        // Act - second sync (should NOT create a new entry)
        var result2 = _tracker.SyncExternalChanges(session.Id);

        // Assert
        Assert.False(result2.HasChanges, "Second sync should report no changes");

        var history = _sessionManager.GetHistory(session.Id);
        var syncEntries = history.Entries.Count(e => e.IsExternalSync);
        Assert.True(syncEntries <= 1, $"Expected at most 1 sync entry, got {syncEntries}");
    }

    [Fact]
    public void SyncExternalChanges_CalledThreeTimes_OnlyCreatesOneWalEntry()
    {
        // Arrange
        var session = _sessionManager.Open(_tempFile);

        // Act
        _tracker.SyncExternalChanges(session.Id);
        var result2 = _tracker.SyncExternalChanges(session.Id);
        var result3 = _tracker.SyncExternalChanges(session.Id);

        // Assert
        Assert.False(result2.HasChanges, "Second sync should report no changes");
        Assert.False(result3.HasChanges, "Third sync should report no changes");
    }

    [Fact]
    public void SyncExternalChanges_AfterFileModified_CreatesNewEntry()
    {
        // Arrange
        var session = _sessionManager.Open(_tempFile);
        _tracker.SyncExternalChanges(session.Id);

        // Modify the external file
        Thread.Sleep(100); // Ensure different timestamp
        ModifyTestDocx(_tempFile, "Modified content");

        // Act
        var result = _tracker.SyncExternalChanges(session.Id);

        // Assert
        Assert.True(result.HasChanges, "Sync after file modification should have changes");
    }

    [Fact]
    public void SyncExternalChanges_AfterModifyThenNoChange_LastSyncHasNoChanges()
    {
        // Arrange
        var session = _sessionManager.Open(_tempFile);

        // First sync (no changes or initial sync)
        _tracker.SyncExternalChanges(session.Id);

        // Modify and sync
        Thread.Sleep(100);
        ModifyTestDocx(_tempFile, "Modified");
        var modifyResult = _tracker.SyncExternalChanges(session.Id);
        Assert.True(modifyResult.HasChanges);

        // Sync again without changes
        var noChangeResult = _tracker.SyncExternalChanges(session.Id);

        // Assert
        Assert.False(noChangeResult.HasChanges, "Sync after modification sync without further changes should have no changes");
    }

    [Fact]
    public void ResolveSession_WithAbsolutePath_ReturnsExistingSession()
    {
        // Arrange
        var session = _sessionManager.Open(_tempFile);
        var absolutePath = Path.GetFullPath(_tempFile);

        // Act
        var resolved = _sessionManager.ResolveSession(absolutePath);

        // Assert
        Assert.Equal(session.Id, resolved.Id);
    }

    [Fact]
    public void ResolveSession_WithSessionId_ReturnsSession()
    {
        // Arrange
        var session = _sessionManager.Open(_tempFile);

        // Act
        var resolved = _sessionManager.ResolveSession(session.Id);

        // Assert
        Assert.Equal(session.Id, resolved.Id);
    }

    [Fact]
    public void ResolveSession_WithNewPath_OpensNewSession()
    {
        // Arrange
        var newFile = Path.Combine(_tempDir, "new.docx");
        CreateTestDocx(newFile, "New file content");

        // Act
        var session = _sessionManager.ResolveSession(newFile);

        // Assert
        Assert.NotNull(session);
        Assert.Equal(Path.GetFullPath(newFile), session.SourcePath);
    }

    [Fact]
    public void ResolveSession_WithSamePath_ReturnsSameSession()
    {
        // Arrange - open via ResolveSession twice
        var session1 = _sessionManager.ResolveSession(_tempFile);
        var session2 = _sessionManager.ResolveSession(_tempFile);

        // Assert
        Assert.Equal(session1.Id, session2.Id);
    }

    [Fact]
    public void Open_WithRelativePath_StoresAbsolutePath()
    {
        // Arrange
        var currentDir = Directory.GetCurrentDirectory();
        Directory.SetCurrentDirectory(_tempDir);

        try
        {
            // Act
            var session = _sessionManager.Open("test.docx");

            // Assert
            Assert.True(Path.IsPathRooted(session.SourcePath));
            // Verify the path ends with the expected filename and is absolute
            Assert.EndsWith("test.docx", session.SourcePath);
            // The actual directory portion may differ due to symlink resolution on macOS
            // (/var -> /private/var), but the path should still be valid
            Assert.True(File.Exists(session.SourcePath), "Stored path should point to existing file");
        }
        finally
        {
            Directory.SetCurrentDirectory(currentDir);
        }
    }

    [Fact]
    public void ResolveSession_WithNonExistentPath_ThrowsKeyNotFound()
    {
        // Arrange
        var nonExistentPath = Path.Combine(_tempDir, "does-not-exist.docx");

        // Act & Assert
        Assert.Throws<KeyNotFoundException>(() => _sessionManager.ResolveSession(nonExistentPath));
    }

    [Fact]
    public void ResolveSession_WithInvalidSessionId_ThrowsKeyNotFound()
    {
        // Act & Assert
        Assert.Throws<KeyNotFoundException>(() => _sessionManager.ResolveSession("invalid123456"));
    }

    #region Helpers

    private static void CreateTestDocx(string path, string content)
    {
        using var ms = new MemoryStream();
        using var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);

        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(
            new Paragraph(new Run(new Text(content)))
        ));

        doc.Save();
        ms.Position = 0;
        File.WriteAllBytes(path, ms.ToArray());
    }

    private static void ModifyTestDocx(string path, string newContent)
    {
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document!.Body!;

        // Clear existing paragraphs
        foreach (var para in body.Elements<Paragraph>().ToList())
        {
            body.RemoveChild(para);
        }

        // Add new content
        body.AppendChild(new Paragraph(new Run(new Text(newContent))));
        doc.Save();
    }

    #endregion

    public void Dispose()
    {
        _tracker.Dispose();

        // Close any open sessions
        foreach (var (id, _) in _sessionManager.List().ToList())
        {
            try { _sessionManager.Close(id); }
            catch { /* ignore */ }
        }

        try { Directory.Delete(_tempDir, true); }
        catch { /* ignore */ }
    }
}
