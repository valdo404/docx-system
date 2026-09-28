using System.Globalization;
using System.Text.Json;

namespace DocxMcp.Layout;

// Absolute layout description (layout.json, version 1).
// Units are typographic points (1/72 in); origin is the top-left corner of the page, y grows downward.
// See docs/layout-to-docx.md for the full schema.

public sealed record LayoutDocument(int Version, IReadOnlyList<LayoutPage> Pages);

public sealed record LayoutPage(double Width, double Height, IReadOnlyList<LayoutItem> Items);

public abstract record LayoutItem;

public sealed record LayoutStroke(string Color, double Width);

/// <summary>
/// A run of text. <see cref="Spacing"/> is an optional extra advance (pt, may be negative) added
/// after each character of the run (w:spacing in w:rPr), used to reproduce the layout engine's
/// justification exactly (e.g. one run per stretched space).
/// </summary>
public sealed record LayoutRun(
    string Text, string? Font, double Size, bool Bold, bool Italic,
    string? Color, bool Underline, string? Link, double? Spacing = null);

public enum LayoutAlign { Left, Right, Center, Justify }

public sealed record LayoutLine(double Baseline, double Height, LayoutAlign Align, IReadOnlyList<LayoutRun> Runs)
{
    /// <summary>Largest run font size on the line (pt); 0 when the line has no run.</summary>
    public double MaxSize => Runs.Count == 0 ? 0 : Runs.Max(r => r.Size);
}

public sealed record LayoutText(double X, double Y, double Width, double Height, IReadOnlyList<LayoutLine> Lines) : LayoutItem;

public sealed record LayoutRect(double X, double Y, double Width, double Height, string? Fill, LayoutStroke? Stroke) : LayoutItem;

public sealed record LayoutLineShape(double X1, double Y1, double X2, double Y2, LayoutStroke Stroke) : LayoutItem;

public sealed record LayoutImage(double X, double Y, double Width, double Height, string Format, byte[] Data) : LayoutItem;

/// <summary>
/// Parses layout.json with <see cref="JsonDocument"/> (reflection-free, NativeAOT safe).
/// Unknown item types are rejected; unknown properties are ignored.
/// </summary>
public static class LayoutParser
{
    public const int SupportedVersion = 1;

    public static LayoutDocument Parse(string json)
    {
        using var doc = JsonDocument.Parse(json, new JsonDocumentOptions
        {
            AllowTrailingCommas = true,
            CommentHandling = JsonCommentHandling.Skip,
        });
        return Parse(doc.RootElement);
    }

    public static LayoutDocument Parse(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object)
            throw new FormatException("layout: root must be an object");

        var version = root.TryGetProperty("version", out var v) ? v.GetInt32() : SupportedVersion;
        if (version != SupportedVersion)
            throw new FormatException($"layout: unsupported version {version} (expected {SupportedVersion})");

        var pages = new List<LayoutPage>();
        var pagesEl = Required(root, "pages", "layout");
        var pi = 0;
        foreach (var p in pagesEl.EnumerateArray())
        {
            var ctx = $"pages[{pi}]";
            var items = new List<LayoutItem>();
            if (p.TryGetProperty("items", out var itemsEl) && itemsEl.ValueKind == JsonValueKind.Array)
            {
                var ii = 0;
                foreach (var it in itemsEl.EnumerateArray())
                    items.Add(ParseItem(it, $"{ctx}.items[{ii++}]"));
            }
            var width = Num(p, "width", ctx);
            var height = Num(p, "height", ctx);
            if (width <= 0 || height <= 0)
                throw new FormatException($"{ctx}: page width and height must be positive");
            pages.Add(new LayoutPage(width, height, items));
            pi++;
        }
        if (pages.Count == 0)
            throw new FormatException("layout: at least one page is required");
        return new LayoutDocument(version, pages);
    }

    private static LayoutItem ParseItem(JsonElement it, string ctx)
    {
        var type = Str(it, "type") ?? throw new FormatException($"{ctx}: missing 'type'");
        return type switch
        {
            "text" => new LayoutText(
                Num(it, "x", ctx), Num(it, "y", ctx), Num(it, "width", ctx), OptNum(it, "height") ?? 0,
                ParseLines(it, ctx)),
            "rect" => new LayoutRect(
                Num(it, "x", ctx), Num(it, "y", ctx), Num(it, "width", ctx), Num(it, "height", ctx),
                Color(Str(it, "fill"), ctx), ParseStroke(it, ctx)),
            "line" => new LayoutLineShape(
                Num(it, "x1", ctx), Num(it, "y1", ctx), Num(it, "x2", ctx), Num(it, "y2", ctx),
                ParseStroke(it, ctx) ?? new LayoutStroke("000000", 1)),
            "image" => new LayoutImage(
                Num(it, "x", ctx), Num(it, "y", ctx), Num(it, "width", ctx), Num(it, "height", ctx),
                (Str(it, "format") ?? "png").ToLowerInvariant(),
                Convert.FromBase64String(Str(it, "data") ?? throw new FormatException($"{ctx}: missing 'data'"))),
            _ => throw new FormatException($"{ctx}: unknown item type '{type}'"),
        };
    }

    private static List<LayoutLine> ParseLines(JsonElement it, string ctx)
    {
        var lines = new List<LayoutLine>();
        if (!it.TryGetProperty("lines", out var linesEl) || linesEl.ValueKind != JsonValueKind.Array)
            return lines;
        var li = 0;
        foreach (var l in linesEl.EnumerateArray())
        {
            var lctx = $"{ctx}.lines[{li++}]";
            var runs = new List<LayoutRun>();
            if (l.TryGetProperty("runs", out var runsEl) && runsEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var r in runsEl.EnumerateArray())
                {
                    runs.Add(new LayoutRun(
                        Str(r, "text") ?? "",
                        Str(r, "font"),
                        OptNum(r, "size") ?? 10,
                        Bool(r, "bold"),
                        Bool(r, "italic"),
                        Color(Str(r, "color"), lctx),
                        Bool(r, "underline"),
                        Str(r, "link"),
                        OptNum(r, "spacing")));
                }
            }
            var align = (Str(l, "align") ?? "left").ToLowerInvariant() switch
            {
                "left" or "start" => LayoutAlign.Left,
                "right" or "end" => LayoutAlign.Right,
                "center" => LayoutAlign.Center,
                "justify" => LayoutAlign.Justify,
                var a => throw new FormatException($"{lctx}: unknown align '{a}'"),
            };
            lines.Add(new LayoutLine(Num(l, "baseline", lctx), Num(l, "height", lctx), align, runs));
        }
        return lines;
    }

    private static LayoutStroke? ParseStroke(JsonElement it, string ctx)
    {
        if (!it.TryGetProperty("stroke", out var s) || s.ValueKind != JsonValueKind.Object)
            return null;
        return new LayoutStroke(Color(Str(s, "color"), ctx) ?? "000000", OptNum(s, "width") ?? 1);
    }

    /// <summary>Normalizes "#RRGGBB" (or "RRGGBB", "#RGB") to upper-case "RRGGBB".</summary>
    internal static string? Color(string? c, string ctx)
    {
        if (string.IsNullOrWhiteSpace(c)) return null;
        var s = c.Trim().TrimStart('#');
        if (s.Length == 3) s = string.Concat(s.Select(ch => new string(ch, 2)));
        if (s.Length == 8) s = s[..6]; // drop alpha (RRGGBBAA)
        if (s.Length != 6 || !int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _))
            throw new FormatException($"{ctx}: invalid color '{c}'");
        return s.ToUpperInvariant();
    }

    private static JsonElement Required(JsonElement e, string name, string ctx) =>
        e.TryGetProperty(name, out var v) && v.ValueKind != JsonValueKind.Null
            ? v : throw new FormatException($"{ctx}: missing '{name}'");

    private static double Num(JsonElement e, string name, string ctx) => Required(e, name, ctx).GetDouble();

    private static double? OptNum(JsonElement e, string name) =>
        e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetDouble() : null;

    private static string? Str(JsonElement e, string name) =>
        e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    private static bool Bool(JsonElement e, string name) =>
        e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.True;
}
