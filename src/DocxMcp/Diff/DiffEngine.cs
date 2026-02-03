using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxMcp.Diff;

/// <summary>
/// Computes structured diffs between two Word documents.
/// Uses content-based matching (LCS algorithm) - does NOT rely on element IDs.
/// Works with any Word document, including those created by Microsoft Word.
/// </summary>
public static class DiffEngine
{
    /// <summary>
    /// Minimum similarity threshold for fuzzy matching (0.0 to 1.0).
    /// </summary>
    public const double DefaultSimilarityThreshold = 0.6;

    /// <summary>
    /// Compare two documents and produce a diff result.
    /// </summary>
    /// <param name="original">The original/baseline document.</param>
    /// <param name="modified">The modified/new document.</param>
    /// <param name="similarityThreshold">Minimum similarity for fuzzy matching (default 0.6).</param>
    /// <returns>A DiffResult containing all detected changes and generated patches.</returns>
    public static DiffResult Compare(
        WordprocessingDocument original,
        WordprocessingDocument modified,
        double similarityThreshold = DefaultSimilarityThreshold)
    {
        var originalBody = original.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Original document has no body.");
        var modifiedBody = modified.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Modified document has no body.");

        // Take snapshots of all top-level content elements
        var originalSnapshots = CaptureSnapshots(originalBody, "/body");
        var modifiedSnapshots = CaptureSnapshots(modifiedBody, "/body");

        // Use LCS-based matching to find correspondences
        var matches = ComputeMatches(originalSnapshots, modifiedSnapshots, similarityThreshold);

        // Generate changes from matches
        var changes = GenerateChanges(originalSnapshots, modifiedSnapshots, matches);

        return new DiffResult { Changes = changes };
    }

    /// <summary>
    /// Compare two documents from byte arrays.
    /// </summary>
    public static DiffResult Compare(
        byte[] originalBytes,
        byte[] modifiedBytes,
        double similarityThreshold = DefaultSimilarityThreshold)
    {
        using var originalStream = new MemoryStream();
        originalStream.Write(originalBytes);
        originalStream.Position = 0;

        using var modifiedStream = new MemoryStream();
        modifiedStream.Write(modifiedBytes);
        modifiedStream.Position = 0;

        using var originalDoc = WordprocessingDocument.Open(originalStream, isEditable: false);
        using var modifiedDoc = WordprocessingDocument.Open(modifiedStream, isEditable: false);

        return Compare(originalDoc, modifiedDoc, similarityThreshold);
    }

    /// <summary>
    /// Compare two documents from file paths.
    /// </summary>
    public static DiffResult Compare(
        string originalPath,
        string modifiedPath,
        double similarityThreshold = DefaultSimilarityThreshold)
    {
        var originalBytes = File.ReadAllBytes(originalPath);
        var modifiedBytes = File.ReadAllBytes(modifiedPath);

        return Compare(originalBytes, modifiedBytes, similarityThreshold);
    }

    /// <summary>
    /// Compare a DocxSession's current state with a file on disk.
    /// Useful for detecting external modifications.
    /// </summary>
    public static DiffResult CompareSessionWithFile(
        DocxSession session,
        string filePath,
        double similarityThreshold = DefaultSimilarityThreshold)
    {
        var sessionBytes = session.ToBytes();
        var fileBytes = File.ReadAllBytes(filePath);

        // Session state = original, file = modified (external changes)
        return Compare(sessionBytes, fileBytes, similarityThreshold);
    }

    /// <summary>
    /// Capture snapshots of all top-level body elements (paragraphs, tables).
    /// </summary>
    private static List<ElementSnapshot> CaptureSnapshots(Body body, string basePath)
    {
        var snapshots = new List<ElementSnapshot>();
        int contentIndex = 0;

        foreach (var element in body.ChildElements)
        {
            // Only track content elements
            if (element is Paragraph or Table)
            {
                snapshots.Add(ElementSnapshot.FromElement(element, contentIndex, basePath));
                contentIndex++;
            }
        }

        return snapshots;
    }

    /// <summary>
    /// Compute matches between original and modified elements using LCS + fuzzy matching.
    /// Returns a dictionary mapping original index to (modified index, match type).
    /// </summary>
    private static MatchResult ComputeMatches(
        List<ElementSnapshot> original,
        List<ElementSnapshot> modified,
        double threshold)
    {
        var result = new MatchResult();

        // Step 1: Find exact matches by fingerprint (identical content)
        var modifiedUsed = new HashSet<int>();
        var exactMatches = new Dictionary<int, int>(); // origIdx -> modIdx

        for (int i = 0; i < original.Count; i++)
        {
            for (int j = 0; j < modified.Count; j++)
            {
                if (modifiedUsed.Contains(j)) continue;

                if (original[i].Fingerprint == modified[j].Fingerprint)
                {
                    exactMatches[i] = j;
                    modifiedUsed.Add(j);
                    break;
                }
            }
        }

        // Step 2: Use LCS on remaining elements to find position-based matches
        var unmatchedOrig = Enumerable.Range(0, original.Count)
            .Where(i => !exactMatches.ContainsKey(i))
            .ToList();
        var unmatchedMod = Enumerable.Range(0, modified.Count)
            .Where(j => !modifiedUsed.Contains(j))
            .ToList();

        var lcsMatches = ComputeLcsMatches(
            unmatchedOrig.Select(i => original[i]).ToList(),
            unmatchedMod.Select(j => modified[j]).ToList(),
            threshold);

        // Map LCS results back to original indices
        var fuzzyMatches = new Dictionary<int, (int modIdx, double similarity)>();
        foreach (var (localOrigIdx, localModIdx, similarity) in lcsMatches)
        {
            var origIdx = unmatchedOrig[localOrigIdx];
            var modIdx = unmatchedMod[localModIdx];
            fuzzyMatches[origIdx] = (modIdx, similarity);
            modifiedUsed.Add(modIdx);
        }

        // Build result
        for (int i = 0; i < original.Count; i++)
        {
            if (exactMatches.TryGetValue(i, out var exactModIdx))
            {
                result.Matches[i] = new ElementMatch
                {
                    OriginalIndex = i,
                    ModifiedIndex = exactModIdx,
                    MatchType = MatchType.Exact,
                    Similarity = 1.0
                };
            }
            else if (fuzzyMatches.TryGetValue(i, out var fuzzy))
            {
                result.Matches[i] = new ElementMatch
                {
                    OriginalIndex = i,
                    ModifiedIndex = fuzzy.modIdx,
                    MatchType = MatchType.Similar,
                    Similarity = fuzzy.similarity
                };
            }
            // Else: no match = deleted
        }

        // Track unmatched modified elements (additions)
        result.UnmatchedModified = Enumerable.Range(0, modified.Count)
            .Where(j => !modifiedUsed.Contains(j))
            .ToList();

        return result;
    }

    /// <summary>
    /// Compute LCS-based matches with similarity scoring.
    /// </summary>
    private static List<(int origIdx, int modIdx, double similarity)> ComputeLcsMatches(
        List<ElementSnapshot> original,
        List<ElementSnapshot> modified,
        double threshold)
    {
        var matches = new List<(int, int, double)>();

        if (original.Count == 0 || modified.Count == 0)
            return matches;

        // Build similarity matrix
        var simMatrix = new double[original.Count, modified.Count];
        for (int i = 0; i < original.Count; i++)
        {
            for (int j = 0; j < modified.Count; j++)
            {
                simMatrix[i, j] = original[i].SimilarityTo(modified[j]);
            }
        }

        // Greedy matching: find best matches above threshold
        var usedOrig = new HashSet<int>();
        var usedMod = new HashSet<int>();

        while (true)
        {
            // Find best remaining match
            double bestSim = 0;
            int bestI = -1, bestJ = -1;

            for (int i = 0; i < original.Count; i++)
            {
                if (usedOrig.Contains(i)) continue;

                for (int j = 0; j < modified.Count; j++)
                {
                    if (usedMod.Contains(j)) continue;

                    if (simMatrix[i, j] > bestSim)
                    {
                        bestSim = simMatrix[i, j];
                        bestI = i;
                        bestJ = j;
                    }
                }
            }

            // Stop if no match above threshold
            if (bestSim < threshold || bestI < 0)
                break;

            matches.Add((bestI, bestJ, bestSim));
            usedOrig.Add(bestI);
            usedMod.Add(bestJ);
        }

        return matches;
    }

    /// <summary>
    /// Generate change list from matches.
    /// </summary>
    private static List<ElementChange> GenerateChanges(
        List<ElementSnapshot> original,
        List<ElementSnapshot> modified,
        MatchResult matches)
    {
        var changes = new List<ElementChange>();

        // Process matched elements
        foreach (var (origIdx, match) in matches.Matches)
        {
            var origSnap = original[origIdx];
            var modSnap = modified[match.ModifiedIndex];

            if (match.MatchType == MatchType.Exact)
            {
                // Check if position changed (moved)
                if (origIdx != match.ModifiedIndex)
                {
                    changes.Add(new ElementChange
                    {
                        ChangeType = ChangeType.Moved,
                        ElementId = origSnap.Fingerprint, // Use fingerprint as identifier
                        ElementType = origSnap.ElementType,
                        OldPath = origSnap.Path,
                        NewPath = modSnap.Path,
                        OldIndex = origIdx,
                        NewIndex = match.ModifiedIndex,
                        OldText = origSnap.Text,
                        NewText = modSnap.Text
                    });
                }
                // Else: no change
            }
            else // Similar match = modification
            {
                changes.Add(new ElementChange
                {
                    ChangeType = ChangeType.Modified,
                    ElementId = origSnap.Fingerprint,
                    ElementType = origSnap.ElementType,
                    OldPath = origSnap.Path,
                    NewPath = modSnap.Path,
                    OldIndex = origIdx,
                    NewIndex = match.ModifiedIndex,
                    OldText = origSnap.Text,
                    NewText = modSnap.Text,
                    OldValue = origSnap.JsonValue,
                    NewValue = CreateValueForPatch(modSnap)
                });
            }
        }

        // Process deletions (unmatched original elements)
        for (int i = 0; i < original.Count; i++)
        {
            if (!matches.Matches.ContainsKey(i))
            {
                var snap = original[i];
                changes.Add(new ElementChange
                {
                    ChangeType = ChangeType.Removed,
                    ElementId = snap.Fingerprint,
                    ElementType = snap.ElementType,
                    OldPath = snap.Path,
                    OldIndex = i,
                    OldText = snap.Text,
                    OldValue = snap.JsonValue
                });
            }
        }

        // Process additions (unmatched modified elements)
        foreach (var modIdx in matches.UnmatchedModified)
        {
            var snap = modified[modIdx];
            changes.Add(new ElementChange
            {
                ChangeType = ChangeType.Added,
                ElementId = snap.Fingerprint,
                ElementType = snap.ElementType,
                NewPath = snap.Path,
                NewIndex = modIdx,
                NewText = snap.Text,
                NewValue = CreateValueForPatch(snap)
            });
        }

        // Sort changes for consistent output
        return changes
            .OrderBy(c => c.ChangeType switch
            {
                ChangeType.Removed => 0,
                ChangeType.Modified => 1,
                ChangeType.Moved => 2,
                ChangeType.Added => 3,
                _ => 4
            })
            .ThenBy(c => c.OldIndex ?? c.NewIndex ?? 0)
            .ToList();
    }

    /// <summary>
    /// Create a JSON value suitable for a patch operation.
    /// </summary>
    private static JsonObject CreateValueForPatch(ElementSnapshot snapshot)
    {
        var value = new JsonObject
        {
            ["type"] = snapshot.ElementType
        };

        // Copy relevant properties from the snapshot's JSON
        foreach (var prop in snapshot.JsonValue)
        {
            if (prop.Key == "type")
                continue;

            value[prop.Key] = prop.Value is not null
                ? JsonNode.Parse(prop.Value.ToJsonString())
                : null;
        }

        return value;
    }

    /// <summary>
    /// Internal result of the matching algorithm.
    /// </summary>
    private sealed class MatchResult
    {
        public Dictionary<int, ElementMatch> Matches { get; } = [];
        public List<int> UnmatchedModified { get; set; } = [];
    }

    /// <summary>
    /// Represents a match between an original and modified element.
    /// </summary>
    private sealed class ElementMatch
    {
        public required int OriginalIndex { get; init; }
        public required int ModifiedIndex { get; init; }
        public required MatchType MatchType { get; init; }
        public required double Similarity { get; init; }
    }

    /// <summary>
    /// Type of match found.
    /// </summary>
    private enum MatchType
    {
        Exact,  // Identical fingerprint
        Similar // Fuzzy match above threshold
    }
}

/// <summary>
/// Extension methods for comparing documents.
/// </summary>
public static class DiffExtensions
{
    /// <summary>
    /// Compare this session with another session.
    /// </summary>
    public static DiffResult CompareTo(this DocxSession session, DocxSession other)
    {
        return DiffEngine.Compare(session.Document, other.Document);
    }

    /// <summary>
    /// Compare this session with a file on disk.
    /// </summary>
    public static DiffResult CompareToFile(this DocxSession session, string filePath)
    {
        return DiffEngine.CompareSessionWithFile(session, filePath);
    }

    /// <summary>
    /// Check if the source file has been modified externally.
    /// </summary>
    public static bool HasExternalChanges(this DocxSession session)
    {
        if (session.SourcePath is null)
            return false;

        if (!File.Exists(session.SourcePath))
            return false;

        var diff = session.CompareToFile(session.SourcePath);
        return diff.HasChanges;
    }
}
