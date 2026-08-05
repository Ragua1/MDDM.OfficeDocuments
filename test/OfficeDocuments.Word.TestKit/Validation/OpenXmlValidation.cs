using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OfficeDocuments.Word.TestKit.Validation;

/// <summary>
/// Schema-validation gate for generated documents.
/// </summary>
/// <remarks>
/// A test that only round-trips a document through this library proves self-consistency, not
/// validity: the file can still violate the OOXML schema and be repaired or rejected by Word. The
/// child-order rules in WordprocessingML are strict and easy to break — <c>w:sectPr</c> has to be the
/// last child of <c>w:body</c>, <c>w:pPr</c> the first child of <c>w:p</c>, and <c>w:rPr</c>'s own
/// children follow a fixed sequence — so every test that produces a complete document should end
/// with an <see cref="AssertValid(string, string[])"/> call.
/// </remarks>
public static class OpenXmlValidation
{
    /// <summary>
    /// Office version whose schema generated documents are validated against.
    /// </summary>
    public const FileFormatVersions TargetFormat = FileFormatVersions.Office2021;

    /// <summary>
    /// Upper bound on how many individual errors are spelled out in an assertion message.
    /// </summary>
    private const int MaxReportedErrors = 25;

    /// <summary>
    /// Asserts that the document stored at <paramref name="filePath"/> is schema-valid.
    /// </summary>
    /// <param name="filePath">Path of the document to validate.</param>
    /// <param name="inheritedDefects">
    /// Substrings of validation-error descriptions to tolerate because they originate in a
    /// pre-existing input document. See <see cref="AssertValid(Stream, string[])"/>.
    /// </param>
    public static void AssertValid(string filePath, params string[] inheritedDefects)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        using var document = WordprocessingDocument.Open(filePath, false);
        AssertValid(document, filePath, inheritedDefects);
    }

    /// <summary>
    /// Asserts that the document held in <paramref name="stream"/> is schema-valid.
    /// The stream position is restored before returning.
    /// </summary>
    /// <param name="stream">Stream holding the document package.</param>
    /// <param name="inheritedDefects">
    /// Substrings of validation-error descriptions to tolerate because they originate in a
    /// pre-existing input document rather than in anything this library produced. Real Word files in
    /// the wild are not always schema-clean, so a test that opens a foreign document must be able to
    /// say "this defect came in with the input" without switching the gate off.
    /// </param>
    public static void AssertValid(Stream stream, params string[] inheritedDefects)
    {
        ArgumentNullException.ThrowIfNull(stream);

        var originalPosition = stream.Position;
        stream.Position = 0;

        try
        {
            using var document = WordprocessingDocument.Open(stream, false);
            AssertValid(document, "<stream>", inheritedDefects);
        }
        finally
        {
            stream.Position = originalPosition;
        }
    }

    /// <summary>
    /// Asserts that an already-open <paramref name="document"/> is schema-valid.
    /// </summary>
    public static void AssertValid(WordprocessingDocument document)
    {
        AssertValid(document, "<document>", []);
    }

    private static void AssertValid(WordprocessingDocument document, string source, string[] inheritedDefects)
    {
        ArgumentNullException.ThrowIfNull(document);

        var validator = new OpenXmlValidator(TargetFormat);
        var errors = validator.Validate(document)
            .Where(error => !IsInherited(error, inheritedDefects))
            .ToList();

        if (errors.Count == 0)
        {
            return;
        }

        Assert.Fail(BuildReport(source, errors));
    }

    private static bool IsInherited(ValidationErrorInfo error, string[] inheritedDefects)
    {
        return inheritedDefects.Any(defect => error.Description.Contains(defect, StringComparison.Ordinal));
    }

    private static string BuildReport(string source, IReadOnlyCollection<ValidationErrorInfo> errors)
    {
        var report = new StringBuilder()
            .Append(errors.Count)
            .Append(errors.Count == 1 ? " schema validation error in '" : " schema validation errors in '")
            .Append(source)
            .Append("' (validated against ")
            .Append(TargetFormat)
            .AppendLine("):");

        foreach (var error in errors.Take(MaxReportedErrors))
        {
            report
                .Append("  [").Append(error.ErrorType).Append("] ").AppendLine(error.Description)
                .Append("      part: ").AppendLine(error.Part?.Uri?.ToString() ?? "<none>")
                .Append("      path: ").AppendLine(error.Path?.XPath ?? "<none>");
        }

        if (errors.Count > MaxReportedErrors)
        {
            report.Append("  ... and ").Append(errors.Count - MaxReportedErrors).AppendLine(" more.");
        }

        return report.ToString();
    }
}
