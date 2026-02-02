using System.Collections.Concurrent;
using System.Text.Json;
using DocxMcp.Persistence;
using Microsoft.Extensions.Logging;

namespace DocxMcp;

/// <summary>
/// Thread-safe manager for document sessions with WAL-based persistence.
/// Sessions survive server restarts via baseline snapshots + write-ahead log replay.
/// Supports undo/redo via WAL cursor + checkpoint replay.
/// </summary>
public sealed class SessionManager
{
    private readonly ConcurrentDictionary<string, DocxSession> _sessions = new();
    private readonly ConcurrentDictionary<string, int> _cursors = new();
    private readonly SessionStore _store;
    private readonly ILogger<SessionManager> _logger;
    private SessionIndexFile _index;
    private readonly object _indexLock = new();
    private readonly int _compactThreshold;
    private readonly int _checkpointInterval;

    public SessionManager(SessionStore store, ILogger<SessionManager> logger)
    {
        _store = store;
        _logger = logger;
        _index = new SessionIndexFile();

        var thresholdEnv = Environment.GetEnvironmentVariable("DOCX_MCP_WAL_COMPACT_THRESHOLD");
        _compactThreshold = int.TryParse(thresholdEnv, out var t) && t > 0 ? t : 50;

        var intervalEnv = Environment.GetEnvironmentVariable("DOCX_MCP_CHECKPOINT_INTERVAL");
        _checkpointInterval = int.TryParse(intervalEnv, out var ci) && ci > 0 ? ci : 10;
    }

    public DocxSession Open(string path)
    {
        var session = DocxSession.Open(path);
        if (!_sessions.TryAdd(session.Id, session))
        {
            session.Dispose();
            throw new InvalidOperationException("Session ID collision — this should not happen.");
        }

        PersistNewSession(session);
        return session;
    }

    public DocxSession Create()
    {
        var session = DocxSession.Create();
        if (!_sessions.TryAdd(session.Id, session))
        {
            session.Dispose();
            throw new InvalidOperationException("Session ID collision — this should not happen.");
        }

        PersistNewSession(session);
        return session;
    }

    public DocxSession Get(string id)
    {
        if (_sessions.TryGetValue(id, out var session))
            return session;
        throw new KeyNotFoundException($"No document session with ID '{id}'.");
    }

    public void Save(string id, string? path = null)
    {
        var session = Get(id);
        session.Save(path);
        // Compact after explicit save — baseline = saved state
        Compact(id);
    }

    public void Close(string id)
    {
        if (_sessions.TryRemove(id, out var session))
        {
            _cursors.TryRemove(id, out _);
            session.Dispose();
            _store.DeleteSession(id);

            lock (_indexLock)
            {
                _index.Sessions.RemoveAll(e => e.Id == id);
                _store.SaveIndex(_index);
            }
        }
        else
        {
            throw new KeyNotFoundException($"No document session with ID '{id}'.");
        }
    }

    public IReadOnlyList<(string Id, string? Path)> List()
    {
        return _sessions.Values
            .Select(s => (s.Id, s.SourcePath))
            .ToList()
            .AsReadOnly();
    }

    // --- WAL operations ---

    /// <summary>
    /// Append a patch to the WAL after a successful mutation.
    /// If the cursor is behind the WAL tip (after undo), truncates future entries first.
    /// Creates checkpoints at interval boundaries.
    /// Triggers compaction if the WAL exceeds the threshold.
    /// </summary>
    public void AppendWal(string id, string patchesJson, string? description = null)
    {
        try
        {
            var cursor = _cursors.GetOrAdd(id, 0);
            var walCount = _store.WalEntryCount(id);

            // If cursor < walCount, we're in an undo state — truncate future
            if (cursor < walCount)
            {
                _store.TruncateWalAt(id, cursor);

                lock (_indexLock)
                {
                    var entry = _index.Sessions.Find(e => e.Id == id);
                    if (entry is not null)
                    {
                        _store.DeleteCheckpointsAfter(id, cursor, entry.CheckpointPositions);
                        entry.CheckpointPositions.RemoveAll(p => p > cursor);
                    }
                }
            }

            // Auto-generate description from patch ops if not provided
            description ??= GenerateDescription(patchesJson);

            _store.AppendWal(id, patchesJson, description);
            var newCursor = cursor + 1;
            _cursors[id] = newCursor;

            // Create checkpoint if crossing an interval boundary
            MaybeCreateCheckpoint(id, newCursor);

            lock (_indexLock)
            {
                var entry = _index.Sessions.Find(e => e.Id == id);
                if (entry is not null)
                {
                    entry.WalCount = _store.WalEntryCount(id);
                    entry.CursorPosition = newCursor;
                    entry.LastModifiedAt = DateTime.UtcNow;
                    _store.SaveIndex(_index);

                    if (entry.WalCount >= _compactThreshold)
                        Compact(id);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to append WAL for session {SessionId}.", id);
        }
    }

    /// <summary>
    /// Create a new baseline snapshot from the current in-memory state and truncate the WAL.
    /// Refuses if redo entries exist unless discardRedoHistory is true.
    /// </summary>
    public void Compact(string id, bool discardRedoHistory = false)
    {
        try
        {
            var cursor = _cursors.GetOrAdd(id, _ => _store.WalEntryCount(id));
            var walCount = _store.WalEntryCount(id);

            if (cursor < walCount && !discardRedoHistory)
            {
                _logger.LogInformation(
                    "Skipping compaction for session {SessionId}: {RedoCount} redo entries exist. Use discardRedoHistory=true to force.",
                    id, walCount - cursor);
                return;
            }

            var session = Get(id);
            var bytes = session.ToBytes();
            _store.PersistBaseline(id, bytes);
            _store.TruncateWal(id);
            _store.DeleteCheckpoints(id);
            _cursors[id] = 0;

            lock (_indexLock)
            {
                var entry = _index.Sessions.Find(e => e.Id == id);
                if (entry is not null)
                {
                    entry.WalCount = 0;
                    entry.CursorPosition = 0;
                    entry.CheckpointPositions.Clear();
                    entry.LastModifiedAt = DateTime.UtcNow;
                    _store.SaveIndex(_index);
                }
            }

            _logger.LogInformation("Compacted session {SessionId}.", id);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to compact session {SessionId}.", id);
        }
    }

    // --- Undo / Redo / JumpTo / History ---

    /// <summary>
    /// Undo N steps by decrementing the cursor and rebuilding from the nearest checkpoint.
    /// </summary>
    public UndoRedoResult Undo(string id, int steps = 1)
    {
        var session = Get(id); // validate session exists
        var cursor = _cursors.GetOrAdd(id, _ => _store.WalEntryCount(id));

        if (cursor <= 0)
            return new UndoRedoResult { Position = 0, Steps = 0, Message = "Already at the beginning. Nothing to undo." };

        var actualSteps = Math.Min(steps, cursor);
        var newCursor = cursor - actualSteps;

        RebuildDocumentAtPosition(id, newCursor);

        return new UndoRedoResult
        {
            Position = newCursor,
            Steps = actualSteps,
            Message = $"Undid {actualSteps} step(s). Now at position {newCursor}."
        };
    }

    /// <summary>
    /// Redo N steps by incrementing the cursor and replaying patches on the current DOM.
    /// </summary>
    public UndoRedoResult Redo(string id, int steps = 1)
    {
        var session = Get(id); // validate session exists
        var cursor = _cursors.GetOrAdd(id, _ => _store.WalEntryCount(id));
        var walCount = _store.WalEntryCount(id);

        if (cursor >= walCount)
            return new UndoRedoResult { Position = cursor, Steps = 0, Message = "Already at the latest state. Nothing to redo." };

        var actualSteps = Math.Min(steps, walCount - cursor);
        var newCursor = cursor + actualSteps;

        // Replay patches [cursor, newCursor) on current DOM (fast, no rebuild)
        var patches = _store.ReadWalRange(id, cursor, newCursor);
        foreach (var patchJson in patches)
        {
            ReplayPatch(session, patchJson);
        }

        _cursors[id] = newCursor;

        lock (_indexLock)
        {
            var entry = _index.Sessions.Find(e => e.Id == id);
            if (entry is not null)
            {
                entry.CursorPosition = newCursor;
                _store.SaveIndex(_index);
            }
        }

        return new UndoRedoResult
        {
            Position = newCursor,
            Steps = actualSteps,
            Message = $"Redid {actualSteps} step(s). Now at position {newCursor}."
        };
    }

    /// <summary>
    /// Jump to an arbitrary WAL position by rebuilding from the nearest checkpoint.
    /// </summary>
    public UndoRedoResult JumpTo(string id, int position)
    {
        var session = Get(id); // validate session exists
        var walCount = _store.WalEntryCount(id);

        if (position < 0)
            position = 0;
        if (position > walCount)
            return new UndoRedoResult
            {
                Position = _cursors.GetOrAdd(id, _ => walCount),
                Steps = 0,
                Message = $"Position {position} is beyond the WAL (max {walCount}). No change."
            };

        var oldCursor = _cursors.GetOrAdd(id, _ => walCount);
        if (position == oldCursor)
            return new UndoRedoResult { Position = position, Steps = 0, Message = $"Already at position {position}." };

        RebuildDocumentAtPosition(id, position);

        var stepsFromOld = Math.Abs(position - oldCursor);
        return new UndoRedoResult
        {
            Position = position,
            Steps = stepsFromOld,
            Message = $"Jumped to position {position}."
        };
    }

    /// <summary>
    /// Get the edit history for a session with metadata.
    /// </summary>
    public HistoryResult GetHistory(string id, int offset = 0, int limit = 20)
    {
        Get(id); // validate session exists
        var walEntries = _store.ReadWalEntries(id);
        var cursor = _cursors.GetOrAdd(id, _ => walEntries.Count);
        var walCount = walEntries.Count;

        List<int> checkpointPositions;
        lock (_indexLock)
        {
            var entry = _index.Sessions.Find(e => e.Id == id);
            checkpointPositions = entry?.CheckpointPositions ?? new List<int>();
        }

        var entries = new List<HistoryEntry>();

        // Include position 0 (baseline) as the first entry
        var startIdx = Math.Max(0, offset);
        var endIdx = Math.Min(walCount + 1, offset + limit); // +1 for baseline

        for (int i = startIdx; i < endIdx; i++)
        {
            if (i == 0)
            {
                entries.Add(new HistoryEntry
                {
                    Position = 0,
                    Timestamp = default,
                    Description = "Baseline (original document)",
                    IsCurrent = cursor == 0,
                    IsCheckpoint = true
                });
            }
            else
            {
                var walIdx = i - 1;
                if (walIdx < walEntries.Count)
                {
                    var we = walEntries[walIdx];
                    entries.Add(new HistoryEntry
                    {
                        Position = i,
                        Timestamp = we.Timestamp,
                        Description = we.Description ?? "",
                        IsCurrent = cursor == i,
                        IsCheckpoint = checkpointPositions.Contains(i)
                    });
                }
            }
        }

        return new HistoryResult
        {
            TotalEntries = walCount + 1, // +1 for baseline
            CursorPosition = cursor,
            CanUndo = cursor > 0,
            CanRedo = cursor < walCount,
            Entries = entries
        };
    }

    /// <summary>
    /// Restore all persisted sessions from disk on startup.
    /// Opens each baseline and replays its WAL up to the cursor position.
    /// </summary>
    public int RestoreSessions()
    {
        _store.EnsureDirectory();
        _index = _store.LoadIndex();
        int restored = 0;
        var stale = new List<string>();

        var maxAgeDaysEnv = Environment.GetEnvironmentVariable("DOCX_MCP_SESSION_MAX_AGE_DAYS");
        var maxAge = TimeSpan.FromDays(int.TryParse(maxAgeDaysEnv, out var d) && d > 0 ? d : 7);
        var cutoff = DateTime.UtcNow - maxAge;

        foreach (var entry in _index.Sessions.ToList())
        {
            // Stale session cleanup
            if (entry.LastModifiedAt < cutoff)
            {
                _logger.LogInformation("Removing stale session {SessionId} (last modified {LastModified}).",
                    entry.Id, entry.LastModifiedAt);
                stale.Add(entry.Id);
                continue;
            }

            try
            {
                var bytes = _store.LoadBaseline(entry.Id);
                var session = DocxSession.FromBytes(bytes, entry.Id, entry.SourcePath);

                // Determine how many WAL entries to replay (up to cursor position)
                var walCount = _store.WalEntryCount(entry.Id);
                var cursorTarget = entry.CursorPosition;

                // Backward compat: old entries with cursor=0 but WAL entries exist
                if (cursorTarget == 0 && walCount > 0 && entry.CheckpointPositions.Count == 0)
                    cursorTarget = walCount;

                var replayCount = Math.Min(cursorTarget, walCount);
                if (replayCount > 0)
                {
                    var patches = _store.ReadWalRange(entry.Id, 0, replayCount);
                    foreach (var patchJson in patches)
                    {
                        try
                        {
                            ReplayPatch(session, patchJson);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to replay WAL entry for session {SessionId}; stopping replay.",
                                entry.Id);
                            break;
                        }
                    }
                }

                if (_sessions.TryAdd(session.Id, session))
                {
                    _cursors[session.Id] = replayCount;
                    restored++;
                }
                else
                    session.Dispose();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to restore session {SessionId}; removing.", entry.Id);
                stale.Add(entry.Id);
            }
        }

        // Clean up stale/corrupt entries
        if (stale.Count > 0)
        {
            lock (_indexLock)
            {
                foreach (var id in stale)
                {
                    _index.Sessions.RemoveAll(e => e.Id == id);
                    _store.DeleteSession(id);
                }
                _store.SaveIndex(_index);
            }
        }

        return restored;
    }

    // --- Private helpers ---

    private void PersistNewSession(DocxSession session)
    {
        try
        {
            var bytes = session.ToBytes();
            _store.PersistBaseline(session.Id, bytes);
            _store.GetOrCreateWal(session.Id); // create empty WAL mapping

            _cursors[session.Id] = 0;

            lock (_indexLock)
            {
                _index.Sessions.Add(new SessionEntry
                {
                    Id = session.Id,
                    SourcePath = session.SourcePath,
                    CreatedAt = DateTime.UtcNow,
                    LastModifiedAt = DateTime.UtcNow,
                    DocxFile = $"{session.Id}.docx",
                    WalCount = 0,
                    CursorPosition = 0
                });
                _store.SaveIndex(_index);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to persist new session {SessionId}.", session.Id);
        }
    }

    /// <summary>
    /// Rebuild the in-memory document at a specific WAL position.
    /// Loads the nearest checkpoint, replays patches to the target position,
    /// and replaces the in-memory session.
    /// </summary>
    private void RebuildDocumentAtPosition(string id, int targetPosition)
    {
        List<int> checkpointPositions;
        lock (_indexLock)
        {
            var indexEntry = _index.Sessions.Find(e => e.Id == id);
            checkpointPositions = indexEntry?.CheckpointPositions ?? new List<int>();
        }

        var (ckptPos, ckptBytes) = _store.LoadNearestCheckpoint(id, targetPosition, checkpointPositions);

        var oldSession = Get(id);
        var newSession = DocxSession.FromBytes(ckptBytes, oldSession.Id, oldSession.SourcePath);

        // Replay patches from checkpoint position to target
        if (targetPosition > ckptPos)
        {
            var patches = _store.ReadWalRange(id, ckptPos, targetPosition);
            foreach (var patchJson in patches)
            {
                try
                {
                    ReplayPatch(newSession, patchJson);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to replay WAL entry during rebuild for session {SessionId}.", id);
                    break;
                }
            }
        }

        // Replace in-memory session
        _sessions[id] = newSession;
        _cursors[id] = targetPosition;
        oldSession.Dispose();

        lock (_indexLock)
        {
            var entry = _index.Sessions.Find(e => e.Id == id);
            if (entry is not null)
            {
                entry.CursorPosition = targetPosition;
                _store.SaveIndex(_index);
            }
        }
    }

    /// <summary>
    /// Create a checkpoint if the new cursor crosses a checkpoint interval boundary.
    /// </summary>
    private void MaybeCreateCheckpoint(string id, int newCursor)
    {
        if (newCursor > 0 && newCursor % _checkpointInterval == 0)
        {
            try
            {
                var session = Get(id);
                var bytes = session.ToBytes();
                _store.PersistCheckpoint(id, newCursor, bytes);

                lock (_indexLock)
                {
                    var entry = _index.Sessions.Find(e => e.Id == id);
                    if (entry is not null && !entry.CheckpointPositions.Contains(newCursor))
                    {
                        entry.CheckpointPositions.Add(newCursor);
                    }
                }

                _logger.LogInformation("Created checkpoint at position {Position} for session {SessionId}.", newCursor, id);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to create checkpoint at position {Position} for session {SessionId}.", newCursor, id);
            }
        }
    }

    /// <summary>
    /// Generate a description string from patch operations.
    /// </summary>
    private static string GenerateDescription(string patchesJson)
    {
        try
        {
            var doc = JsonDocument.Parse(patchesJson);
            if (doc.RootElement.ValueKind != JsonValueKind.Array)
                return "patch";

            var ops = new List<string>();
            foreach (var patch in doc.RootElement.EnumerateArray())
            {
                var op = patch.TryGetProperty("op", out var opEl) ? opEl.GetString() : null;
                var path = patch.TryGetProperty("path", out var pathEl) ? pathEl.GetString() : null;
                if (op is not null)
                {
                    var shortPath = path is not null && path.Length > 30
                        ? path[..30] + "..."
                        : path;
                    ops.Add(shortPath is not null ? $"{op} {shortPath}" : op);
                }
            }

            return ops.Count > 0 ? string.Join(", ", ops) : "patch";
        }
        catch
        {
            return "patch";
        }
    }

    /// <summary>
    /// Replay a single patch operation against a session's document.
    /// Uses the same logic as PatchTool.ApplyPatch but without MCP tool wiring.
    /// </summary>
    private static void ReplayPatch(DocxSession session, string patchesJson)
    {
        var patchArray = JsonDocument.Parse(patchesJson).RootElement;
        if (patchArray.ValueKind != JsonValueKind.Array)
            return;

        var wpDoc = session.Document;
        var mainPart = wpDoc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no MainDocumentPart.");

        foreach (var patch in patchArray.EnumerateArray())
        {
            var op = patch.GetProperty("op").GetString()?.ToLowerInvariant();
            switch (op)
            {
                case "add":
                    Tools.PatchTool.ReplayAdd(patch, wpDoc, mainPart);
                    break;
                case "replace":
                    Tools.PatchTool.ReplayReplace(patch, wpDoc, mainPart);
                    break;
                case "remove":
                    Tools.PatchTool.ReplayRemove(patch, wpDoc);
                    break;
                case "move":
                    Tools.PatchTool.ReplayMove(patch, wpDoc);
                    break;
                case "copy":
                    Tools.PatchTool.ReplayCopy(patch, wpDoc);
                    break;
                case "replace_text":
                    Tools.PatchTool.ReplayReplaceText(patch, wpDoc);
                    break;
                case "remove_column":
                    Tools.PatchTool.ReplayRemoveColumn(patch, wpDoc);
                    break;
            }
        }
    }
}
