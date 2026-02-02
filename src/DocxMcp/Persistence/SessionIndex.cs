using System.Text.Json.Serialization;

namespace DocxMcp.Persistence;

public sealed class SessionIndexFile
{
    public int Version { get; set; } = 1;
    public List<SessionEntry> Sessions { get; set; } = new();
}

public sealed class SessionEntry
{
    public string Id { get; set; } = "";
    public string? SourcePath { get; set; }
    public DateTime CreatedAt { get; set; }
    public DateTime LastModifiedAt { get; set; }
    public string DocxFile { get; set; } = "";
    public int WalCount { get; set; }
    public int CursorPosition { get; set; } = -1;
    public List<int> CheckpointPositions { get; set; } = new();
}

[JsonSerializable(typeof(SessionIndexFile))]
[JsonSerializable(typeof(SessionEntry))]
[JsonSerializable(typeof(List<SessionEntry>))]
[JsonSerializable(typeof(List<int>))]
[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.SnakeCaseLower,
    WriteIndented = true)]
internal partial class SessionJsonContext : JsonSerializerContext { }
