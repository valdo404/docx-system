using System.Globalization;

namespace DocxMcp.Layout;

/// <summary>
/// Where Word puts the baseline inside a line whose spacing is "exactly <c>lineHeight</c>".
///
/// Word does not document this; the model is empirical and this class is the ONLY place that
/// encodes it, so it can be re-calibrated:
///
///   offset(lineHeight, size) = ratio × lineHeight        (ratio = 0.8 by default)
///
/// Calibration (Word for Mac 16, compatibility mode 15, PDF export, 2026-09): single-line text
/// boxes in Helvetica, Arial, Times New Roman and Georgia at 8/10/12/24 pt with exact line
/// heights from 0.5× to 3× the font size all put the baseline at 0.80 × lineHeight below the
/// line top, within the 0.24 pt resolution of Word's PDF output — independently of the font
/// and of the font size, including lines smaller than the font size.
///
/// The ratio can be tuned with the <c>--baseline-ratio</c> CLI flag or the
/// <c>DOCX_LAYOUT_BASELINE_RATIO</c> environment variable. <c>fontSizePt</c> is kept in the
/// signature so a font-dependent model can be plugged in without touching callers.
/// </summary>
public sealed class BaselineModel
{
    public const string EnvVar = "DOCX_LAYOUT_BASELINE_RATIO";
    public const double DefaultRatio = 0.8;

    public double Ratio { get; }

    public BaselineModel(double ratio = DefaultRatio)
    {
        if (ratio <= 0 || ratio >= 2 || double.IsNaN(ratio))
            throw new ArgumentOutOfRangeException(nameof(ratio), "baseline ratio must be in (0, 2)");
        Ratio = ratio;
    }

    public static BaselineModel Default { get; } = new();

    /// <summary>Model from <see cref="EnvVar"/> when set, else the default.</summary>
    public static BaselineModel FromEnvironment()
    {
        var env = Environment.GetEnvironmentVariable(EnvVar);
        return string.IsNullOrWhiteSpace(env)
            ? Default
            : new BaselineModel(double.Parse(env, CultureInfo.InvariantCulture));
    }

    /// <summary>Baseline offset (pt) from the top of an exact-height line.</summary>
    public double BaselineOffset(double lineHeight, double fontSizePt) => Ratio * lineHeight;

    /// <summary>
    /// Smallest exact line height whose baseline lands <paramref name="targetOffset"/> below the
    /// line top (inverse of <see cref="BaselineOffset"/>, by bisection so any monotone model works).
    /// Returns null when the target is below the minimum reachable offset.
    /// </summary>
    public double? SolveLineHeight(double targetOffset, double fontSizePt)
    {
        const double min = 0.05;
        if (BaselineOffset(min, fontSizePt) > targetOffset + 1e-6)
            return null;
        double lo = min, hi = Math.Max(1, targetOffset * 4 + fontSizePt * 4);
        while (BaselineOffset(hi, fontSizePt) < targetOffset) hi *= 2;
        for (var i = 0; i < 60; i++)
        {
            var mid = (lo + hi) / 2;
            if (BaselineOffset(mid, fontSizePt) < targetOffset) lo = mid; else hi = mid;
        }
        return hi;
    }
}
