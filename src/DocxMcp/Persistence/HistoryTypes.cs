namespace DocxMcp.Persistence;

public sealed class UndoRedoResult
{
    public int Position { get; set; }
    public int Steps { get; set; }
    public string Message { get; set; } = "";
}

public sealed class HistoryEntry
{
    public int Position { get; set; }
    public DateTime Timestamp { get; set; }
    public string Description { get; set; } = "";
    public bool IsCurrent { get; set; }
    public bool IsCheckpoint { get; set; }
}

public sealed class HistoryResult
{
    public int TotalEntries { get; set; }
    public int CursorPosition { get; set; }
    public bool CanUndo { get; set; }
    public bool CanRedo { get; set; }
    public List<HistoryEntry> Entries { get; set; } = new();
}
