namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Covers creating, closing, and reopening a document.
/// </summary>
public class DocumentLifecycleTest : WordTestBase
{
    [Fact]
    public void CreateDocument_EmptyBody_ProducesSchemaValidFile()
    {
        using var stream = WriteAndValidate(_ => { });

        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void CreateDocument_ToFilePath_WritesFile()
    {
        var filePath = GetFilepath("lifecycle.docx");

        using (var document = CreateDocument(filePath))
        {
            document.GetBody().AddParagraph("Persisted through a file path");
        }

        Assert.True(File.Exists(filePath));
        OpenXmlValidation.AssertValid(filePath);
    }

    /// <summary>
    /// Regression test: the documented usage pattern is <c>using</c> plus an explicit
    /// <c>Close()</c>, which used to save an already-disposed package and throw.
    /// </summary>
    [Fact]
    public void Close_CalledBeforeDispose_DoesNotThrow()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Closed explicitly");
            document.Close();
        }

        OpenXmlValidation.AssertValid(stream);
    }

    [Fact]
    public void Close_CalledTwice_DoesNotThrow()
    {
        using var stream = new MemoryStream();
        var document = CreateDocument(stream);
        document.GetBody().AddParagraph("Closed twice");

        document.Close();
        document.Close();

        OpenXmlValidation.AssertValid(stream);
    }

    [Fact]
    public void GetBody_AfterClose_Throws()
    {
        using var stream = new MemoryStream();
        var document = CreateDocument(stream);
        document.Close();

        Assert.Throws<ObjectDisposedException>(() => document.GetBody());
    }

    [Fact]
    public void GetBody_CalledTwice_ReturnsSameBody()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);

        Assert.Same(document.GetBody(), document.GetBody());
    }

    [Fact]
    public void OpenDocument_ReadOnly_DoesNotChangeTheFile()
    {
        var filePath = GetFilepath("read-only.docx");
        using (var document = CreateDocument(filePath))
        {
            document.GetBody().AddParagraph("Original content");
        }

        var originalBytes = File.ReadAllBytes(filePath);

        using (var document = OpenDocument(filePath, isEditable: false))
        {
            Assert.Equal("Original content", document.GetBody().GetAllTexts());
        }

        Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
    }

    [Fact]
    public void OpenDocument_ThenAppend_KeepsExistingContent()
    {
        var filePath = GetFilepath("append.docx");
        using (var document = CreateDocument(filePath))
        {
            document.GetBody().AddParagraph("First");
        }

        using (var document = OpenDocument(filePath))
        {
            document.GetBody().AddParagraph("Second");
        }

        using var reopened = OpenDocument(filePath, isEditable: false);
        Assert.Equal("First\nSecond", reopened.GetBody().GetAllTexts());
        OpenXmlValidation.AssertValid(filePath);
    }
}
