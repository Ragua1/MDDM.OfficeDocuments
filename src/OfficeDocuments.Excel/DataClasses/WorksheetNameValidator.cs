namespace OfficeDocuments.Excel.DataClasses;

/// <summary>
/// Excel's rules for what a worksheet may be called.
/// <para>
/// Without this the library writes whatever it is given, and the result is a workbook Excel offers
/// to repair — the name attribute is constrained in the schema itself, by both a maximum length
/// and a character pattern, so an over-long or `/`-bearing name produces a file that is invalid
/// rather than merely unusual. Every comparable library has shipped this bug at least once, and
/// the failure is silent at the point of the mistake and loud in the user's face much later.
/// </para>
/// </summary>
internal static class WorksheetNameValidator
{
    /// <summary>Excel's limit, and the schema's <c>MaxLength</c> on the attribute.</summary>
    public const int MaxLength = 31;

    /// <summary>
    /// Excel reserves this name for a shared workbook's change history and refuses to create it.
    /// </summary>
    private const string ReservedName = "History";

    /// <summary>
    /// The characters Excel forbids anywhere in a sheet name. They are the ones that would be
    /// ambiguous inside an A1-style reference — <c>Sheet1!A1</c>, <c>[Book]Sheet</c>, <c>'a b'!A1</c>.
    /// </summary>
    private static readonly char[] ForbiddenCharacters = [':', '\\', '/', '?', '*', '[', ']'];

    /// <exception cref="ArgumentException">The name is not one Excel will accept.</exception>
    public static void Validate(string? name, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("A worksheet name cannot be empty or whitespace.", parameterName);
        }

        if (name.Length > MaxLength)
        {
            throw new ArgumentException(
                $"A worksheet name cannot be longer than {MaxLength} characters; '{name}' is {name.Length}.",
                parameterName);
        }

        var forbiddenIndex = name.IndexOfAny(ForbiddenCharacters);
        if (forbiddenIndex >= 0)
        {
            throw new ArgumentException(
                $"A worksheet name cannot contain '{name[forbiddenIndex]}'. Forbidden characters are "
                + $"{string.Join(' ', ForbiddenCharacters)}.",
                parameterName);
        }

        // A leading apostrophe is rejected by the schema pattern; a trailing one is rejected by
        // Excel. Both would break the quoting of 'Sheet Name'!A1 references.
        if (name[0] == '\'' || name[^1] == '\'')
        {
            throw new ArgumentException("A worksheet name cannot start or end with an apostrophe.", parameterName);
        }

        if (string.Equals(name, ReservedName, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException($"'{ReservedName}' is reserved by Excel and cannot be used as a worksheet name.", parameterName);
        }

        XmlText.EnsureRepresentable(name, parameterName, "The worksheet name");
    }
}
