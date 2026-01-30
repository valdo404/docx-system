using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.Json;
using Xunit;

namespace DocxMcp.Tests;

public class PatchLimitTests : IDisposable
{
    private readonly DocxSession _session;
    private readonly SessionManager _sessions;

    public PatchLimitTests()
    {
        _sessions = new SessionManager();
        _session = _sessions.Create();

        var body = _session.GetBody();
        body.AppendChild(new Paragraph(new Run(new Text("Content"))));
    }

    [Fact]
    public void TenPatchesAreAccepted()
    {
        var patches = new List<object>();
        for (int i = 0; i < 10; i++)
        {
            patches.Add(new
            {
                op = "add",
                path = "/body/children/0",
                value = new { type = "paragraph", text = $"Added {i}" }
            });
        }

        var json = JsonSerializer.Serialize(patches);
        var result = DocxMcp.Tools.PatchTool.ApplyPatch(_sessions, _session.Id, json);

        Assert.Contains("Applied 10 patch(es) successfully", result);
    }

    [Fact]
    public void ElevenPatchesAreRejected()
    {
        var patches = new List<object>();
        for (int i = 0; i < 11; i++)
        {
            patches.Add(new
            {
                op = "add",
                path = "/body/children/0",
                value = new { type = "paragraph", text = $"Added {i}" }
            });
        }

        var json = JsonSerializer.Serialize(patches);
        var result = DocxMcp.Tools.PatchTool.ApplyPatch(_sessions, _session.Id, json);

        Assert.Contains("Error: Too many operations (11)", result);
        Assert.Contains("Maximum is 10", result);
    }

    [Fact]
    public void OnePatchIsAccepted()
    {
        var json = """[{"op": "add", "path": "/body/children/0", "value": {"type": "paragraph", "text": "Hello"}}]""";
        var result = DocxMcp.Tools.PatchTool.ApplyPatch(_sessions, _session.Id, json);

        Assert.Contains("Applied 1 patch(es) successfully", result);
    }

    [Fact]
    public void EmptyPatchArrayIsAccepted()
    {
        var result = DocxMcp.Tools.PatchTool.ApplyPatch(_sessions, _session.Id, "[]");

        Assert.Contains("Applied 0 patch(es) successfully", result);
    }

    public void Dispose()
    {
        _sessions.Close(_session.Id);
    }
}
