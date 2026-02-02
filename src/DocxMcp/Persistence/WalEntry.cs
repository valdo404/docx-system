using System.Text.Json.Serialization;

namespace DocxMcp.Persistence;

public sealed class WalEntry
{
    public string Patches { get; set; } = "";
    public DateTime Timestamp { get; set; }
    public string? Description { get; set; }
}

[JsonSerializable(typeof(WalEntry))]
[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.SnakeCaseLower)]
internal partial class WalJsonContext : JsonSerializerContext { }
