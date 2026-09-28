using System.ComponentModel;
using DocxMcp.Layout;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace DocxMcp.Tools;

/// <summary>
/// Builds a brand new .docx from an absolute layout (e.g. produced by Typst).
/// Stateless: no session is opened; see docs/layout-to-docx.md for the schema.
/// </summary>
[McpServerToolType]
public sealed class LayoutTools
{
    [McpServerTool(Name = "document_from_layout"), Description(
        "Create a new .docx from an ABSOLUTE layout description (layout.json version 1: pages of " +
        "text/rect/line/image items positioned in points from the top-left of the page). " +
        "Each page becomes a Word section of the exact page size; every item becomes an anchored " +
        "drawing object (text box, shape or picture) at its exact position.\n\n" +
        "If output_path is given, the file is written there and a summary is returned; " +
        "otherwise the .docx bytes are returned base64-encoded. No session is created.")]
    public static string DocumentFromLayout(
        [Description("The layout JSON (version 1).")] string layout_json,
        [Description("Optional path of the .docx to write. If omitted, returns base64.")] string? output_path = null,
        [Description("Optional baseline position ratio inside exact-height lines (default 0.8).")] double? baseline_ratio = null)
    {
        try
        {
            var layout = LayoutParser.Parse(layout_json);
            var options = new LayoutDocxOptions
            {
                Baseline = baseline_ratio is { } r ? new BaselineModel(r) : BaselineModel.FromEnvironment(),
            };
            using var ms = new MemoryStream();
            LayoutDocxWriter.Write(layout, ms, options);
            var items = layout.Pages.Sum(p => p.Items.Count);

            if (string.IsNullOrWhiteSpace(output_path))
                return Convert.ToBase64String(ms.ToArray());

            File.WriteAllBytes(output_path, ms.ToArray());
            return $"Wrote {output_path}: {layout.Pages.Count} page(s), {items} item(s), {ms.Length} bytes.";
        }
        catch (McpException) { throw; }
        catch (Exception ex) { throw new McpException($"document_from_layout: {ex.Message}", ex); }
    }
}
