using System.Text.Json;
using DocxMcp.Persistence;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace DocxMcp.Tests;

public class SessionStoreTests : IDisposable
{
    private readonly string _tempDir;
    private readonly SessionStore _store;

    public SessionStoreTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "docx-mcp-tests", Guid.NewGuid().ToString("N"));
        _store = new SessionStore(NullLogger<SessionStore>.Instance, _tempDir);
    }

    public void Dispose()
    {
        _store.Dispose();
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    // --- Index tests ---

    [Fact]
    public void LoadIndex_NoFile_ReturnsEmpty()
    {
        var index = _store.LoadIndex();
        Assert.Equal(1, index.Version);
        Assert.Empty(index.Sessions);
    }

    [Fact]
    public void SaveAndLoadIndex_RoundTrips()
    {
        var index = new SessionIndexFile
        {
            Sessions = new()
            {
                new SessionEntry
                {
                    Id = "abc123",
                    SourcePath = "/tmp/test.docx",
                    CreatedAt = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    LastModifiedAt = new DateTime(2026, 1, 2, 0, 0, 0, DateTimeKind.Utc),
                    DocxFile = "abc123.docx",
                    WalCount = 5
                }
            }
        };

        _store.SaveIndex(index);
        var loaded = _store.LoadIndex();

        Assert.Single(loaded.Sessions);
        var entry = loaded.Sessions[0];
        Assert.Equal("abc123", entry.Id);
        Assert.Equal("/tmp/test.docx", entry.SourcePath);
        Assert.Equal("abc123.docx", entry.DocxFile);
        Assert.Equal(5, entry.WalCount);
    }

    [Fact]
    public void SaveIndex_MultipleSessions_RoundTrips()
    {
        var index = new SessionIndexFile
        {
            Sessions = new()
            {
                new SessionEntry { Id = "aaa", DocxFile = "aaa.docx" },
                new SessionEntry { Id = "bbb", DocxFile = "bbb.docx" },
                new SessionEntry { Id = "ccc", DocxFile = "ccc.docx" },
            }
        };

        _store.SaveIndex(index);
        var loaded = _store.LoadIndex();

        Assert.Equal(3, loaded.Sessions.Count);
        Assert.Equal("aaa", loaded.Sessions[0].Id);
        Assert.Equal("bbb", loaded.Sessions[1].Id);
        Assert.Equal("ccc", loaded.Sessions[2].Id);
    }

    [Fact]
    public void LoadIndex_NullSourcePath_RoundTrips()
    {
        var index = new SessionIndexFile
        {
            Sessions = new()
            {
                new SessionEntry { Id = "x", SourcePath = null, DocxFile = "x.docx" }
            }
        };

        _store.SaveIndex(index);
        var loaded = _store.LoadIndex();

        Assert.Null(loaded.Sessions[0].SourcePath);
    }

    [Fact]
    public void LoadIndex_CorruptJson_ReturnsEmpty()
    {
        _store.EnsureDirectory();
        File.WriteAllText(Path.Combine(_tempDir, "index.json"), "not valid json {{{");

        var index = _store.LoadIndex();
        Assert.Empty(index.Sessions);
    }

    [Fact]
    public void SaveIndex_CreatesDirectoryIfMissing()
    {
        Assert.False(Directory.Exists(_tempDir));
        _store.SaveIndex(new SessionIndexFile());
        Assert.True(Directory.Exists(_tempDir));
        Assert.True(File.Exists(Path.Combine(_tempDir, "index.json")));
    }

    // --- Baseline tests ---

    [Fact]
    public void PersistAndLoadBaseline_RoundTrips()
    {
        var data = new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x01, 0x02, 0x03 };
        _store.PersistBaseline("sess1", data);

        var loaded = _store.LoadBaseline("sess1");
        Assert.Equal(data, loaded);
    }

    [Fact]
    public void PersistBaseline_LargeData_RoundTrips()
    {
        var data = new byte[500_000];
        new Random(42).NextBytes(data);

        _store.PersistBaseline("large", data);
        var loaded = _store.LoadBaseline("large");

        Assert.Equal(data.Length, loaded.Length);
        Assert.Equal(data, loaded);
    }

    [Fact]
    public void PersistBaseline_Overwrite_ReplacesOldData()
    {
        _store.PersistBaseline("s1", new byte[] { 1, 2, 3 });
        _store.PersistBaseline("s1", new byte[] { 4, 5, 6, 7 });

        var loaded = _store.LoadBaseline("s1");
        Assert.Equal(new byte[] { 4, 5, 6, 7 }, loaded);
    }

    [Fact]
    public void LoadBaseline_MissingFile_Throws()
    {
        Assert.ThrowsAny<IOException>(() => _store.LoadBaseline("nonexistent"));
    }

    [Fact]
    public void PersistBaseline_CreatesDirectoryIfMissing()
    {
        Assert.False(Directory.Exists(_tempDir));
        _store.PersistBaseline("s1", new byte[] { 0xFF });
        Assert.True(File.Exists(_store.BaselinePath("s1")));
    }

    // --- DeleteSession tests ---

    [Fact]
    public void DeleteSession_RemovesBothFiles()
    {
        _store.PersistBaseline("del1", new byte[] { 1, 2 });
        _store.GetOrCreateWal("del1");

        Assert.True(File.Exists(_store.BaselinePath("del1")));
        Assert.True(File.Exists(_store.WalPath("del1")));

        _store.DeleteSession("del1");

        Assert.False(File.Exists(_store.BaselinePath("del1")));
        Assert.False(File.Exists(_store.WalPath("del1")));
    }

    [Fact]
    public void DeleteSession_NonExistent_DoesNotThrow()
    {
        _store.DeleteSession("ghost"); // should not throw
    }

    [Fact]
    public void DeleteSession_AlsoRemovesCheckpoints()
    {
        _store.PersistBaseline("ck1", new byte[] { 1, 2 });
        _store.PersistCheckpoint("ck1", 10, new byte[] { 3, 4 });
        _store.PersistCheckpoint("ck1", 20, new byte[] { 5, 6 });
        _store.GetOrCreateWal("ck1");

        _store.DeleteSession("ck1");

        Assert.False(File.Exists(_store.CheckpointPath("ck1", 10)));
        Assert.False(File.Exists(_store.CheckpointPath("ck1", 20)));
    }

    // --- WAL integration with store ---

    [Fact]
    public void AppendWal_AndReadWal_RoundTrips()
    {
        _store.AppendWal("w1", "[{\"op\":\"add\"}]");
        _store.AppendWal("w1", "[{\"op\":\"remove\"}]");

        var patches = _store.ReadWal("w1");
        Assert.Equal(2, patches.Count);
        Assert.Equal("[{\"op\":\"add\"}]", patches[0]);
        Assert.Equal("[{\"op\":\"remove\"}]", patches[1]);
    }

    [Fact]
    public void WalEntryCount_TracksCorrectly()
    {
        Assert.Equal(0, _store.WalEntryCount("w2"));

        _store.AppendWal("w2", "[{\"op\":\"add\"}]");
        Assert.Equal(1, _store.WalEntryCount("w2"));

        _store.AppendWal("w2", "[{\"op\":\"remove\"}]");
        Assert.Equal(2, _store.WalEntryCount("w2"));
    }

    [Fact]
    public void TruncateWal_ClearsEntries()
    {
        _store.AppendWal("w3", "[{\"op\":\"add\"}]");
        _store.TruncateWal("w3");

        Assert.Equal(0, _store.WalEntryCount("w3"));
        Assert.Empty(_store.ReadWal("w3"));
    }

    [Fact]
    public void ReadWal_NoWalFile_ReturnsEmpty()
    {
        // GetOrCreateWal creates the file, but ReadWal on a fresh store should handle missing file
        var patches = _store.ReadWal("nowal");
        Assert.Empty(patches);
    }

    // --- JSON serialization tests ---

    [Fact]
    public void SessionJsonContext_ProducesSnakeCaseKeys()
    {
        var entry = new SessionEntry
        {
            Id = "test",
            SourcePath = "/path",
            CreatedAt = DateTime.UtcNow,
            LastModifiedAt = DateTime.UtcNow,
            DocxFile = "test.docx",
            WalCount = 3
        };

        var index = new SessionIndexFile { Sessions = new() { entry } };
        var json = JsonSerializer.Serialize(index, SessionJsonContext.Default.SessionIndexFile);

        Assert.Contains("\"source_path\"", json);
        Assert.Contains("\"created_at\"", json);
        Assert.Contains("\"last_modified_at\"", json);
        Assert.Contains("\"docx_file\"", json);
        Assert.Contains("\"wal_count\"", json);
        Assert.DoesNotContain("\"SourcePath\"", json);
        Assert.DoesNotContain("\"WalCount\"", json);
    }

    [Fact]
    public void SessionJsonContext_IncludesCursorAndCheckpoints()
    {
        var entry = new SessionEntry
        {
            Id = "test",
            DocxFile = "test.docx",
            CursorPosition = 5,
            CheckpointPositions = new() { 10, 20 }
        };

        var index = new SessionIndexFile { Sessions = new() { entry } };
        var json = JsonSerializer.Serialize(index, SessionJsonContext.Default.SessionIndexFile);

        Assert.Contains("\"cursor_position\"", json);
        Assert.Contains("\"checkpoint_positions\"", json);

        var loaded = JsonSerializer.Deserialize(json, SessionJsonContext.Default.SessionIndexFile);
        Assert.NotNull(loaded);
        Assert.Equal(5, loaded!.Sessions[0].CursorPosition);
        Assert.Equal(new List<int> { 10, 20 }, loaded.Sessions[0].CheckpointPositions);
    }

    [Fact]
    public void WalJsonContext_ProducesSnakeCaseKeys()
    {
        var entry = new WalEntry { Patches = "[{\"op\":\"add\"}]" };
        var json = JsonSerializer.Serialize(entry, WalJsonContext.Default.WalEntry);

        Assert.Contains("\"patches\"", json);
        Assert.DoesNotContain("\"Patches\"", json);

        var deserialized = JsonSerializer.Deserialize(json, WalJsonContext.Default.WalEntry);
        Assert.NotNull(deserialized);
        Assert.Equal("[{\"op\":\"add\"}]", deserialized!.Patches);
    }

    [Fact]
    public void WalJsonContext_IncludesTimestampAndDescription()
    {
        var entry = new WalEntry
        {
            Patches = "[]",
            Timestamp = new DateTime(2026, 1, 15, 12, 0, 0, DateTimeKind.Utc),
            Description = "add /body/paragraph[0]"
        };
        var json = JsonSerializer.Serialize(entry, WalJsonContext.Default.WalEntry);

        Assert.Contains("\"timestamp\"", json);
        Assert.Contains("\"description\"", json);

        var deserialized = JsonSerializer.Deserialize(json, WalJsonContext.Default.WalEntry);
        Assert.NotNull(deserialized);
        Assert.Equal("add /body/paragraph[0]", deserialized!.Description);
    }

    // --- Path helpers ---

    [Fact]
    public void BaselinePath_IncludesSessionId()
    {
        var path = _store.BaselinePath("abc123");
        Assert.EndsWith("abc123.docx", path);
        Assert.StartsWith(_tempDir, path);
    }

    [Fact]
    public void WalPath_IncludesSessionId()
    {
        var path = _store.WalPath("abc123");
        Assert.EndsWith("abc123.wal", path);
        Assert.StartsWith(_tempDir, path);
    }

    // --- Checkpoint tests ---

    [Fact]
    public void CheckpointPath_Format()
    {
        var path = _store.CheckpointPath("sess1", 10);
        Assert.EndsWith("sess1.ckpt.10.docx", path);
        Assert.StartsWith(_tempDir, path);
    }

    [Fact]
    public void PersistAndLoadCheckpoint_RoundTrips()
    {
        var data = new byte[] { 0xAA, 0xBB, 0xCC };
        _store.PersistCheckpoint("ck1", 10, data);

        Assert.True(File.Exists(_store.CheckpointPath("ck1", 10)));

        // Load via LoadNearestCheckpoint
        var (pos, bytes) = _store.LoadNearestCheckpoint("ck1", 10, new List<int> { 10 });
        Assert.Equal(10, pos);
        Assert.Equal(data, bytes);
    }

    [Fact]
    public void LoadNearestCheckpoint_SelectsNearest()
    {
        _store.PersistBaseline("ck2", new byte[] { 0x01 });
        _store.PersistCheckpoint("ck2", 10, new byte[] { 0x0A });
        _store.PersistCheckpoint("ck2", 20, new byte[] { 0x14 });

        // Target 15: nearest <= 15 is 10
        var (pos, bytes) = _store.LoadNearestCheckpoint("ck2", 15, new List<int> { 10, 20 });
        Assert.Equal(10, pos);
        Assert.Equal(new byte[] { 0x0A }, bytes);

        // Target 25: nearest <= 25 is 20
        (pos, bytes) = _store.LoadNearestCheckpoint("ck2", 25, new List<int> { 10, 20 });
        Assert.Equal(20, pos);
        Assert.Equal(new byte[] { 0x14 }, bytes);
    }

    [Fact]
    public void LoadNearestCheckpoint_FallsBackToBaseline()
    {
        _store.PersistBaseline("ck3", new byte[] { 0xFF });
        _store.PersistCheckpoint("ck3", 10, new byte[] { 0x0A });

        // Target 5: no checkpoint <= 5 (only 10), fallback to baseline
        var (pos, bytes) = _store.LoadNearestCheckpoint("ck3", 5, new List<int> { 10 });
        Assert.Equal(0, pos);
        Assert.Equal(new byte[] { 0xFF }, bytes);
    }

    [Fact]
    public void DeleteCheckpoints_RemovesAll()
    {
        _store.PersistCheckpoint("ck4", 10, new byte[] { 1 });
        _store.PersistCheckpoint("ck4", 20, new byte[] { 2 });

        Assert.True(File.Exists(_store.CheckpointPath("ck4", 10)));
        Assert.True(File.Exists(_store.CheckpointPath("ck4", 20)));

        _store.DeleteCheckpoints("ck4");

        Assert.False(File.Exists(_store.CheckpointPath("ck4", 10)));
        Assert.False(File.Exists(_store.CheckpointPath("ck4", 20)));
    }

    [Fact]
    public void DeleteCheckpointsAfter_RemovesOnlyLater()
    {
        _store.PersistCheckpoint("ck5", 10, new byte[] { 1 });
        _store.PersistCheckpoint("ck5", 20, new byte[] { 2 });
        _store.PersistCheckpoint("ck5", 30, new byte[] { 3 });

        _store.DeleteCheckpointsAfter("ck5", 15, new List<int> { 10, 20, 30 });

        Assert.True(File.Exists(_store.CheckpointPath("ck5", 10)));
        Assert.False(File.Exists(_store.CheckpointPath("ck5", 20)));
        Assert.False(File.Exists(_store.CheckpointPath("ck5", 30)));
    }

    // --- ReadWalRange tests ---

    [Fact]
    public void ReadWalRange_ReturnsSubset()
    {
        _store.AppendWal("wr1", "[{\"op\":\"add\"}]");
        _store.AppendWal("wr1", "[{\"op\":\"remove\"}]");
        _store.AppendWal("wr1", "[{\"op\":\"replace\"}]");

        var range = _store.ReadWalRange("wr1", 1, 3);
        Assert.Equal(2, range.Count);
        Assert.Equal("[{\"op\":\"remove\"}]", range[0]);
        Assert.Equal("[{\"op\":\"replace\"}]", range[1]);
    }

    // --- TruncateWalAt tests ---

    [Fact]
    public void TruncateWalAt_KeepsFirstN()
    {
        _store.AppendWal("tw1", "[{\"op\":\"add\"}]");
        _store.AppendWal("tw1", "[{\"op\":\"remove\"}]");
        _store.AppendWal("tw1", "[{\"op\":\"replace\"}]");

        _store.TruncateWalAt("tw1", 2);

        Assert.Equal(2, _store.WalEntryCount("tw1"));
        var patches = _store.ReadWal("tw1");
        Assert.Equal(2, patches.Count);
        Assert.Equal("[{\"op\":\"add\"}]", patches[0]);
        Assert.Equal("[{\"op\":\"remove\"}]", patches[1]);
    }

    // --- AppendWal with description ---

    [Fact]
    public void AppendWal_WithDescription_RoundTrips()
    {
        _store.AppendWal("wd1", "[{\"op\":\"add\"}]", "add paragraph");

        var entries = _store.ReadWalEntries("wd1");
        Assert.Single(entries);
        Assert.Equal("[{\"op\":\"add\"}]", entries[0].Patches);
        Assert.Equal("add paragraph", entries[0].Description);
        Assert.True(entries[0].Timestamp > DateTime.MinValue);
    }

    // --- ReadWalEntries tests ---

    [Fact]
    public void ReadWalEntries_ReturnsFullMetadata()
    {
        _store.AppendWal("we1", "[{\"op\":\"add\"}]", "first op");
        _store.AppendWal("we1", "[{\"op\":\"remove\"}]", "second op");

        var entries = _store.ReadWalEntries("we1");
        Assert.Equal(2, entries.Count);
        Assert.Equal("first op", entries[0].Description);
        Assert.Equal("second op", entries[1].Description);
        Assert.Equal("[{\"op\":\"add\"}]", entries[0].Patches);
        Assert.Equal("[{\"op\":\"remove\"}]", entries[1].Patches);
    }
}
