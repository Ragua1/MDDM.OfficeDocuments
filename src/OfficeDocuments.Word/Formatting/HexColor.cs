namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Normalizes the colour strings that WordprocessingML attributes accept.
/// </summary>
internal static class HexColor
{
    /// <summary>
    /// Value that tells Word to pick the colour itself.
    /// </summary>
    internal const string Automatic = "auto";

    /// <summary>
    /// Validates <paramref name="value"/> and returns it in the <c>RRGGBB</c> form Word expects.
    /// </summary>
    /// <param name="value">A 6-digit hex colour, optionally <c>#</c>-prefixed, or <c>auto</c>.</param>
    /// <param name="parameterName">Name reported in the exception.</param>
    /// <exception cref="ArgumentException">The value is not a colour Word can store.</exception>
    internal static string Normalize(string value, string parameterName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(value, parameterName);

        var candidate = value.Trim();
        if (string.Equals(candidate, Automatic, StringComparison.OrdinalIgnoreCase))
        {
            return Automatic;
        }

        if (candidate.StartsWith('#'))
        {
            candidate = candidate[1..];
        }

        if (candidate.Length == 8)
        {
            // Refusing beats silently discarding the alpha channel: a caller who passes ARGB is
            // expecting transparency that WordprocessingML simply cannot express.
            throw new ArgumentException(
                $"Word colors are 6 hex digits (RRGGBB) and have no alpha channel; '{value}' has 8. Drop the leading alpha pair.",
                parameterName);
        }

        if (candidate.Length != 6 || !IsHex(candidate))
        {
            throw new ArgumentException(
                $"'{value}' is not a valid color. Expected 6 hex digits (RRGGBB), optionally prefixed with '#', or '{Automatic}'.",
                parameterName);
        }

        return candidate.ToUpperInvariant();
    }

    private static bool IsHex(string value)
    {
        foreach (var character in value)
        {
            if (!Uri.IsHexDigit(character))
            {
                return false;
            }
        }

        return true;
    }
}
