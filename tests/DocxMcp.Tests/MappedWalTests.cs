using DocxMcp.Persistence;
using Xunit;

namespace DocxMcp.Tests;

public class MappedWalTests : IDisposable
{
    private readonly string _tempDir;

    public MappedWalTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "docx-mcp-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    private string WalPath(string name = "test") => Path.Combine(_tempDir, $"{name}.wal");

    [Fact]
    public void NewWal_IsEmpty()
    {
        using var wal = new MappedWal(WalPath());
        Assert.Empty(wal.ReadAll());
        Assert.Equal(0, wal.EntryCount);
    }

    [Fact]
    public void Append_SingleEntry_CanBeRead()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("line one");

        var lines = wal.ReadAll();
        Assert.Single(lines);
        Assert.Equal("line one", lines[0]);
        Assert.Equal(1, wal.EntryCount);
    }

    [Fact]
    public void Append_MultipleEntries_PreservesOrder()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("first");
        wal.Append("second");
        wal.Append("third");

        var lines = wal.ReadAll();
        Assert.Equal(3, lines.Count);
        Assert.Equal("first", lines[0]);
        Assert.Equal("second", lines[1]);
        Assert.Equal("third", lines[2]);
        Assert.Equal(3, wal.EntryCount);
    }

    [Fact]
    public void Truncate_ClearsAllEntries()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("data");
        wal.Append("more data");

        wal.Truncate();

        Assert.Empty(wal.ReadAll());
        Assert.Equal(0, wal.EntryCount);
    }

    [Fact]
    public void Truncate_ThenAppend_Works()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("old");
        wal.Truncate();
        wal.Append("new");

        var lines = wal.ReadAll();
        Assert.Single(lines);
        Assert.Equal("new", lines[0]);
    }

    [Fact]
    public void Persistence_SurvivesReopen()
    {
        var path = WalPath();

        using (var wal = new MappedWal(path))
        {
            wal.Append("persisted line 1");
            wal.Append("persisted line 2");
        }

        using (var wal2 = new MappedWal(path))
        {
            var lines = wal2.ReadAll();
            Assert.Equal(2, lines.Count);
            Assert.Equal("persisted line 1", lines[0]);
            Assert.Equal("persisted line 2", lines[1]);
        }
    }

    [Fact]
    public void Persistence_TruncatedWal_ReopensEmpty()
    {
        var path = WalPath();

        using (var wal = new MappedWal(path))
        {
            wal.Append("will be truncated");
            wal.Truncate();
        }

        using (var wal2 = new MappedWal(path))
        {
            Assert.Empty(wal2.ReadAll());
        }
    }

    [Fact]
    public void Grow_HandlesLargeAppends()
    {
        using var wal = new MappedWal(WalPath());

        // Append enough data to exceed the initial 1MB capacity
        var largeLine = new string('x', 50_000);
        for (int i = 0; i < 25; i++)
            wal.Append(largeLine);

        var lines = wal.ReadAll();
        Assert.Equal(25, lines.Count);
        Assert.All(lines, l => Assert.Equal(largeLine, l));
    }

    [Fact]
    public void Append_Utf8Content_RoundTrips()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("{\"text\":\"héllo wörld 日本語\"}");

        var lines = wal.ReadAll();
        Assert.Single(lines);
        Assert.Equal("{\"text\":\"héllo wörld 日本語\"}", lines[0]);
    }

    [Fact]
    public void EntryCount_MatchesAppendCount()
    {
        using var wal = new MappedWal(WalPath());
        Assert.Equal(0, wal.EntryCount);

        for (int i = 1; i <= 10; i++)
        {
            wal.Append($"entry {i}");
            Assert.Equal(i, wal.EntryCount);
        }
    }

    // --- ReadRange tests ---

    [Fact]
    public void ReadRange_Subset_ReturnsCorrectEntries()
    {
        using var wal = new MappedWal(WalPath());
        for (int i = 0; i < 5; i++)
            wal.Append($"line {i}");

        var range = wal.ReadRange(1, 4);
        Assert.Equal(3, range.Count);
        Assert.Equal("line 1", range[0]);
        Assert.Equal("line 2", range[1]);
        Assert.Equal("line 3", range[2]);
    }

    [Fact]
    public void ReadRange_FullRange_ReturnsAll()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("a");
        wal.Append("b");
        wal.Append("c");

        var range = wal.ReadRange(0, 3);
        Assert.Equal(3, range.Count);
        Assert.Equal("a", range[0]);
        Assert.Equal("b", range[1]);
        Assert.Equal("c", range[2]);
    }

    [Fact]
    public void ReadRange_EmptyRange_ReturnsEmpty()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("data");

        Assert.Empty(wal.ReadRange(1, 1));
        Assert.Empty(wal.ReadRange(2, 1));
    }

    [Fact]
    public void ReadRange_OutOfBounds_ClampsSafely()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("a");
        wal.Append("b");

        var range = wal.ReadRange(-1, 100);
        Assert.Equal(2, range.Count);
        Assert.Equal("a", range[0]);
        Assert.Equal("b", range[1]);
    }

    [Fact]
    public void ReadRange_OnEmptyWal_ReturnsEmpty()
    {
        using var wal = new MappedWal(WalPath());
        Assert.Empty(wal.ReadRange(0, 10));
    }

    // --- ReadEntry tests ---

    [Fact]
    public void ReadEntry_ByIndex_ReturnsCorrect()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("alpha");
        wal.Append("beta");
        wal.Append("gamma");

        Assert.Equal("alpha", wal.ReadEntry(0));
        Assert.Equal("beta", wal.ReadEntry(1));
        Assert.Equal("gamma", wal.ReadEntry(2));
    }

    [Fact]
    public void ReadEntry_OutOfRange_Throws()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("only one");

        Assert.Throws<ArgumentOutOfRangeException>(() => wal.ReadEntry(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => wal.ReadEntry(1));
        Assert.Throws<ArgumentOutOfRangeException>(() => wal.ReadEntry(100));
    }

    // --- TruncateAt tests ---

    [Fact]
    public void TruncateAt_KeepsFirstN()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("a");
        wal.Append("b");
        wal.Append("c");
        wal.Append("d");

        wal.TruncateAt(2);

        Assert.Equal(2, wal.EntryCount);
        var lines = wal.ReadAll();
        Assert.Equal("a", lines[0]);
        Assert.Equal("b", lines[1]);
    }

    [Fact]
    public void TruncateAt_Zero_ClearsAll()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("a");
        wal.Append("b");

        wal.TruncateAt(0);

        Assert.Equal(0, wal.EntryCount);
        Assert.Empty(wal.ReadAll());
    }

    [Fact]
    public void TruncateAt_BeyondCount_NoOp()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("a");
        wal.Append("b");

        wal.TruncateAt(10);

        Assert.Equal(2, wal.EntryCount);
        Assert.Equal(2, wal.ReadAll().Count);
    }

    [Fact]
    public void TruncateAt_ThenAppend_Works()
    {
        using var wal = new MappedWal(WalPath());
        wal.Append("a");
        wal.Append("b");
        wal.Append("c");

        wal.TruncateAt(1);
        wal.Append("new b");

        Assert.Equal(2, wal.EntryCount);
        var lines = wal.ReadAll();
        Assert.Equal("a", lines[0]);
        Assert.Equal("new b", lines[1]);
    }

    [Fact]
    public void TruncateAt_Persistence_SurvivesReopen()
    {
        var path = WalPath();

        using (var wal = new MappedWal(path))
        {
            wal.Append("a");
            wal.Append("b");
            wal.Append("c");
            wal.TruncateAt(2);
        }

        using (var wal2 = new MappedWal(path))
        {
            Assert.Equal(2, wal2.EntryCount);
            var lines = wal2.ReadAll();
            Assert.Equal("a", lines[0]);
            Assert.Equal("b", lines[1]);
        }
    }

    [Fact]
    public void OffsetIndex_RebuiltOnReopen()
    {
        var path = WalPath();

        using (var wal = new MappedWal(path))
        {
            wal.Append("line 0");
            wal.Append("line 1");
            wal.Append("line 2");
        }

        using (var wal2 = new MappedWal(path))
        {
            // Verify random access works after reopen (offset index rebuilt)
            Assert.Equal("line 0", wal2.ReadEntry(0));
            Assert.Equal("line 1", wal2.ReadEntry(1));
            Assert.Equal("line 2", wal2.ReadEntry(2));
            Assert.Equal(3, wal2.EntryCount);

            var range = wal2.ReadRange(1, 3);
            Assert.Equal(2, range.Count);
            Assert.Equal("line 1", range[0]);
            Assert.Equal("line 2", range[1]);
        }
    }
}
