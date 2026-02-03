using System.ComponentModel;
using System.Text.Json;
using DocxMcp.ExternalChanges;
using ModelContextProtocol.Server;

namespace DocxMcp.Tools;

/// <summary>
/// MCP tool for handling external document changes.
/// Single unified tool that detects, displays, and acknowledges external modifications.
/// </summary>
[McpServerToolType]
public static class ExternalChangeTools
{
    /// <summary>
    /// Check for external changes, get details, and optionally acknowledge them.
    /// This is the single tool for all external change operations.
    /// </summary>
    [McpServerTool(Name = "get_external_changes"), Description(
        "Check if the source file has been modified externally and get change details.\n\n" +
        "This tool:\n" +
        "1. Detects if the source file was modified outside this session\n" +
        "2. Shows detailed diff (what was added, removed, modified, moved)\n" +
        "3. Can acknowledge changes to allow continued editing\n\n" +
        "IMPORTANT: If external changes are detected, you MUST acknowledge them " +
        "(set acknowledge=true) before you can use apply_patch on this document.")]
    public static ExternalChangeResult GetExternalChanges(
        ExternalChangeTracker tracker,
        [Description("Session ID to check for external changes")]
        string doc_id,
        [Description("Set to true to acknowledge the changes and allow editing to continue")]
        bool acknowledge = false)
    {
        // First check for any already-detected pending changes
        var pending = tracker.GetLatestUnacknowledgedChange(doc_id);

        // If no pending, check for new changes
        if (pending is null)
        {
            pending = tracker.CheckForChanges(doc_id);
        }

        // No changes detected
        if (pending is null)
        {
            return new ExternalChangeResult
            {
                HasChanges = false,
                Message = "No external changes detected. The document is in sync with the source file.",
                CanEdit = true
            };
        }

        // Acknowledge if requested
        if (acknowledge)
        {
            tracker.AcknowledgeChange(doc_id, pending.Id);

            return new ExternalChangeResult
            {
                HasChanges = true,
                Acknowledged = true,
                CanEdit = true,
                ChangeId = pending.Id,
                DetectedAt = pending.DetectedAt,
                SourcePath = pending.SourcePath,
                Summary = new ChangeSummary
                {
                    TotalChanges = pending.Summary.TotalChanges,
                    Added = pending.Summary.Added,
                    Removed = pending.Summary.Removed,
                    Modified = pending.Summary.Modified,
                    Moved = pending.Summary.Moved
                },
                Changes = pending.Changes.Select(c => new ChangeDetail
                {
                    Type = c.ChangeType,
                    ElementType = c.ElementType,
                    Description = c.Description,
                    OldText = c.OldText,
                    NewText = c.NewText
                }).ToList(),
                Message = $"External changes acknowledged. You may now continue editing.\n\n" +
                          $"Summary: {pending.Summary.TotalChanges} change(s) were made externally:\n" +
                          $"  • {pending.Summary.Added} added\n" +
                          $"  • {pending.Summary.Removed} removed\n" +
                          $"  • {pending.Summary.Modified} modified\n" +
                          $"  • {pending.Summary.Moved} moved"
            };
        }

        // Return details without acknowledging
        return new ExternalChangeResult
        {
            HasChanges = true,
            Acknowledged = false,
            CanEdit = false,
            ChangeId = pending.Id,
            DetectedAt = pending.DetectedAt,
            SourcePath = pending.SourcePath,
            Summary = new ChangeSummary
            {
                TotalChanges = pending.Summary.TotalChanges,
                Added = pending.Summary.Added,
                Removed = pending.Summary.Removed,
                Modified = pending.Summary.Modified,
                Moved = pending.Summary.Moved
            },
            Changes = pending.Changes.Select(c => new ChangeDetail
            {
                Type = c.ChangeType,
                ElementType = c.ElementType,
                Description = c.Description,
                OldText = c.OldText,
                NewText = c.NewText
            }).ToList(),
            Patches = pending.Patches.Select(p => p.ToJsonString()).ToList(),
            Message = BuildChangeMessage(pending)
        };
    }

    private static string BuildChangeMessage(ExternalChangePatch patch)
    {
        var lines = new List<string>
        {
            "⚠️ EXTERNAL CHANGES DETECTED",
            "",
            $"The file '{Path.GetFileName(patch.SourcePath)}' was modified externally.",
            $"Detected at: {patch.DetectedAt:yyyy-MM-dd HH:mm:ss UTC}",
            "",
            "## Summary",
            $"  • Added: {patch.Summary.Added}",
            $"  • Removed: {patch.Summary.Removed}",
            $"  • Modified: {patch.Summary.Modified}",
            $"  • Moved: {patch.Summary.Moved}",
            $"  • Total: {patch.Summary.TotalChanges}",
            ""
        };

        if (patch.Changes.Count > 0)
        {
            lines.Add("## Changes");
            foreach (var change in patch.Changes.Take(15))
            {
                lines.Add($"  • {change.Description}");
            }
            if (patch.Changes.Count > 15)
            {
                lines.Add($"  • ... and {patch.Changes.Count - 15} more");
            }
            lines.Add("");
        }

        lines.Add("## Action Required");
        lines.Add("Call `get_external_changes` with `acknowledge=true` to continue editing.");

        return string.Join("\n", lines);
    }
}

#region Result Types

public sealed class ExternalChangeResult
{
    /// <summary>Whether external changes were detected.</summary>
    public required bool HasChanges { get; init; }

    /// <summary>Whether the changes have been acknowledged.</summary>
    public bool Acknowledged { get; init; }

    /// <summary>Whether editing is allowed (true if no changes or acknowledged).</summary>
    public required bool CanEdit { get; init; }

    /// <summary>Unique identifier for this change event.</summary>
    public string? ChangeId { get; init; }

    /// <summary>When the change was detected.</summary>
    public DateTime? DetectedAt { get; init; }

    /// <summary>Path to the source file.</summary>
    public string? SourcePath { get; init; }

    /// <summary>Summary counts of changes.</summary>
    public ChangeSummary? Summary { get; init; }

    /// <summary>List of individual changes.</summary>
    public List<ChangeDetail>? Changes { get; init; }

    /// <summary>Generated patches (for reference).</summary>
    public List<string>? Patches { get; init; }

    /// <summary>Human-readable message.</summary>
    public required string Message { get; init; }
}

public sealed class ChangeSummary
{
    public int TotalChanges { get; init; }
    public int Added { get; init; }
    public int Removed { get; init; }
    public int Modified { get; init; }
    public int Moved { get; init; }
}

public sealed class ChangeDetail
{
    public required string Type { get; init; }
    public required string ElementType { get; init; }
    public required string Description { get; init; }
    public string? OldText { get; init; }
    public string? NewText { get; init; }
}

#endregion
