using System.Globalization;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Converts the point-based measurements this library exposes into the units WordprocessingML
/// attributes actually store.
/// </summary>
/// <remarks>
/// WordprocessingML measures the same quantity in different units depending on the attribute:
/// font size is in half-points, spacing and indentation in twentieths of a point ("twips"), and
/// line spacing in twips again but with its own rule multiplier. Making callers know which is which
/// is the sort of accidental complexity this library exists to remove, so the public surface takes
/// points everywhere and converts here.
/// </remarks>
internal static class Measure
{
    /// <summary>
    /// Largest font size Word accepts, in points.
    /// </summary>
    internal const double MaxFontSizeInPoints = 1638d;

    /// <summary>
    /// Twentieths of a point in a point.
    /// </summary>
    private const double TwipsPerPoint = 20d;

    /// <summary>
    /// Converts a font size in points to the half-points <c>w:sz</c> stores.
    /// </summary>
    internal static string FontSizeToHalfPoints(double points)
    {
        return ToInvariant((int)Math.Round(points * 2d, MidpointRounding.AwayFromZero));
    }

    /// <summary>
    /// Converts a length in points to the twips <c>w:spacing</c> and <c>w:ind</c> store.
    /// </summary>
    internal static string PointsToTwips(double points)
    {
        return ToInvariant((int)Math.Round(points * TwipsPerPoint, MidpointRounding.AwayFromZero));
    }

    /// <summary>
    /// Converts a multiple of single line spacing to the twips <c>w:spacing/@w:line</c> stores,
    /// where one line is 240 twips.
    /// </summary>
    internal static string LineSpacingToTwips(double lines)
    {
        return ToInvariant((int)Math.Round(lines * 240d, MidpointRounding.AwayFromZero));
    }

    /// <summary>
    /// Validates a font size expressed in points.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The size is outside what Word can store.</exception>
    internal static double ValidateFontSize(double points, string parameterName)
    {
        if (double.IsNaN(points) || points <= 0d || points > MaxFontSizeInPoints)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                $"Font size must be greater than 0 and at most {MaxFontSizeInPoints} points.");
        }

        return points;
    }

    /// <summary>
    /// Validates a length expressed in points, rejecting the non-finite values that would produce
    /// a malformed attribute rather than an error.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not a usable measurement.</exception>
    internal static double ValidateLength(double points, string parameterName, bool allowNegative = false)
    {
        if (double.IsNaN(points) || double.IsInfinity(points) || (!allowNegative && points < 0d))
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                allowNegative
                    ? "Measurement must be a finite number of points."
                    : "Measurement must be a finite, non-negative number of points.");
        }

        return points;
    }

    private static string ToInvariant(int value) => value.ToString(CultureInfo.InvariantCulture);
}
