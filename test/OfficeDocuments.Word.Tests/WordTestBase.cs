using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Word.Interfaces;
using OfficeDocuments.Word.TestKit;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Shared plumbing for the Word tests.
/// </summary>
/// <remarks>
/// The default authoring target is a <see cref="MemoryStream"/>: a document in memory is faster,
/// leaves nothing behind, and cannot collide between parallel test classes. The file-based helpers
/// exist for the tests that specifically exercise the path-based entry points.
/// </remarks>
public abstract class WordTestBase : IDisposable
{
    protected IWordprocessing CreateDocument(Stream stream) => new Wordprocessing(stream, createNew: true);

    protected IWordprocessing CreateDocument(string filePath) => new Wordprocessing(filePath, createNew: true);

    protected IWordprocessing OpenDocument(Stream stream, bool isEditable = true) => new Wordprocessing(stream, createNew: false, isEditable);

    protected IWordprocessing OpenDocument(string filePath, bool isEditable = true) => new Wordprocessing(filePath, createNew: false, isEditable);

    protected string GetFilepath(string filename) => TempWorkspace.GetFilepath(this, filename);

    /// <summary>
    /// Writes a document through <paramref name="author"/>, asserts it is schema-valid, and returns
    /// the package for further inspection.
    /// </summary>
    /// <remarks>
    /// Validating here rather than in each test means a schema-order regression fails the suite even
    /// when the test that provoked it was only checking something else.
    /// </remarks>
    protected MemoryStream WriteAndValidate(Action<IBody> author)
    {
        var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            author(document.GetBody());
        }

        OpenXmlValidation.AssertValid(stream);

        return stream;
    }

    /// <summary>
    /// Reads the raw main-document XML of a written package, for assertions about the markup itself.
    /// </summary>
    protected static string ReadMainDocumentXml(MemoryStream stream)
    {
        stream.Position = 0;
        using var document = WordprocessingDocument.Open(stream, false);

        return document.MainDocumentPart?.Document?.OuterXml ?? string.Empty;
    }

    /// <summary>
    /// Projects something out of a written package's document element.
    /// </summary>
    /// <remarks>
    /// Preferred over matching against <see cref="ReadMainDocumentXml"/> when the point of the test is
    /// which elements and values are present, not how the serializer spells them. Both
    /// <c>w:val="0"</c> and <c>w:val="false"</c> mean the same thing to Word, so a string assertion
    /// would be pinning the SDK's formatting rather than the library's behaviour.
    /// </remarks>
    protected static T ReadDocumentElement<T>(MemoryStream stream, Func<WordLib.Document, T> read)
    {
        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var documentElement = package.MainDocumentPart?.Document
            ?? throw new InvalidOperationException("The package has no main document part.");

        return read(documentElement);
    }

    /// <summary>
    /// Reads the raw styles XML of a written package, or an empty string when there is no styles part.
    /// </summary>
    protected static string ReadStylesXml(MemoryStream stream)
    {
        stream.Position = 0;
        using var document = WordprocessingDocument.Open(stream, false);

        return document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.OuterXml ?? string.Empty;
    }

    public void Dispose()
    {
        TempWorkspace.Cleanup(this);
        GC.SuppressFinalize(this);
    }
}
