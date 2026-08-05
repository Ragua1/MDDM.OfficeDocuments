using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Covers bullet and numbered lists, including the numbering definitions they depend on.
/// </summary>
public class ListTest : WordTestBase
{
    /// <summary>
    /// A list item carries only a pointer into the numbering part. Without the definition it renders
    /// as an ordinary paragraph, so the definition is the feature.
    /// </summary>
    [Fact]
    public void AddListItem_CreatesTheNumberingDefinitionItReferences()
    {
        using var stream = WriteAndValidate(body => body.AddListItem("First"));

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var numbering = package.MainDocumentPart!.NumberingDefinitionsPart?.Numbering;

        Assert.NotNull(numbering);

        var instance = Assert.Single(numbering.Elements<WordLib.NumberingInstance>());
        var abstractNumbering = Assert.Single(numbering.Elements<WordLib.AbstractNum>());

        Assert.Equal(abstractNumbering.AbstractNumberId?.Value, instance.AbstractNumId?.Val?.Value);

        var reference = package.MainDocumentPart.Document!.Descendants<WordLib.NumberingId>().Single();
        Assert.Equal(instance.NumberID?.Value, reference.Val?.Value);
    }

    /// <summary>
    /// <c>CT_Numbering</c> is <c>numPicBullet*, abstractNum*, num*</c>, so the abstract definitions have
    /// to precede the concrete instances.
    /// </summary>
    [Fact]
    public void NumberingPart_KeepsAbstractDefinitionsBeforeInstances()
    {
        using var stream = WriteAndValidate(body =>
        {
            body.AddListItem("bullet", ListStyle.Bullet);
            body.AddListItem("number", ListStyle.Number);
        });

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var children = package.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.ChildElements
            .Select(child => child.LocalName)
            .ToList();

        var lastAbstract = children.LastIndexOf("abstractNum");
        var firstInstance = children.IndexOf("num");

        Assert.True(lastAbstract < firstInstance, $"Expected every abstractNum before the first num, got: {string.Join(", ", children)}");
    }

    /// <summary>
    /// A definition declares all nine levels, because a level that is referenced but not defined falls
    /// back to the document default rather than to anything list-shaped.
    /// </summary>
    [Fact]
    public void NumberingDefinition_DeclaresEveryLevel()
    {
        using var stream = WriteAndValidate(body => body.AddListItem("First"));

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var abstractNumbering = package.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<WordLib.AbstractNum>()
            .Single();

        Assert.Equal(9, abstractNumbering.Elements<WordLib.Level>().Count());
    }

    [Fact]
    public void AddListItem_Bullet_UsesTheBulletNumberFormat()
    {
        using var stream = WriteAndValidate(body => body.AddListItem("First", ListStyle.Bullet));

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var firstLevel = package.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<WordLib.AbstractNum>()
            .Single()
            .Elements<WordLib.Level>()
            .First();

        Assert.Equal(WordLib.NumberFormatValues.Bullet, firstLevel.NumberingFormat?.Val?.Value);
    }

    [Fact]
    public void AddListItem_Number_UsesTheDecimalNumberFormat()
    {
        using var stream = WriteAndValidate(body => body.AddListItem("First", ListStyle.Number));

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var firstLevel = package.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<WordLib.AbstractNum>()
            .Single()
            .Elements<WordLib.Level>()
            .First();

        Assert.Equal(WordLib.NumberFormatValues.Decimal, firstLevel.NumberingFormat?.Val?.Value);
        Assert.Equal("%1.", firstLevel.LevelText?.Val?.Value);
    }

    /// <summary>
    /// Two lists of the same kind should share one definition; a definition per item would bloat the
    /// numbering part and, worse, restart the numbering on every item.
    /// </summary>
    [Fact]
    public void AddListItem_RepeatedForOneStyle_ReusesTheSameNumbering()
    {
        using var stream = WriteAndValidate(body =>
        {
            body.AddListItem("First");
            body.AddListItem("Second");
            body.AddListItem("Third");
        });

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var numbering = package.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;

        Assert.Single(numbering.Elements<WordLib.AbstractNum>());
        Assert.Single(numbering.Elements<WordLib.NumberingInstance>());

        var referencedIds = package.MainDocumentPart.Document!
            .Descendants<WordLib.NumberingId>()
            .Select(id => id.Val?.Value)
            .Distinct()
            .ToList();

        Assert.Single(referencedIds);
    }

    [Fact]
    public void AddListItem_ForBothStyles_CreatesOneNumberingEach()
    {
        using var stream = WriteAndValidate(body =>
        {
            body.AddListItem("bullet", ListStyle.Bullet);
            body.AddListItem("number", ListStyle.Number);
        });

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var numbering = package.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;

        Assert.Equal(2, numbering.Elements<WordLib.AbstractNum>().Count());
        Assert.Equal(2, numbering.Elements<WordLib.NumberingInstance>().Count());
    }

    [Fact]
    public void AddListItem_WithLevel_RoundTripsTheLevelAndStyle()
    {
        using var stream = WriteAndValidate(body =>
        {
            body.AddListItem("Top", ListStyle.Number);
            body.AddListItem("Nested", ListStyle.Number, level: 1);
        });

        using var document = OpenDocument(stream, isEditable: false);
        var paragraphs = document.GetBody().Paragraphs;

        Assert.Equal(ListStyle.Number, paragraphs[0].Format.ListStyle);
        Assert.Equal(0, paragraphs[0].Format.ListLevel);
        Assert.Equal(ListStyle.Number, paragraphs[1].Format.ListStyle);
        Assert.Equal(1, paragraphs[1].Format.ListLevel);
    }

    /// <summary>
    /// The style is resolved by reading the numbering definition, not by remembering what was written,
    /// so it also works for a list in a document this library did not author.
    /// </summary>
    [Fact]
    public void ListStyle_IsResolvedFromTheDocumentAfterReopening()
    {
        var filePath = GetFilepath("list-reopen.docx");
        using (var document = CreateDocument(filePath))
        {
            document.GetBody().AddListItem("Bulleted", ListStyle.Bullet);
        }

        using var reopened = OpenDocument(filePath, isEditable: false);

        Assert.Equal(ListStyle.Bullet, reopened.GetBody().Paragraphs[0].Format.ListStyle);
    }

    /// <summary>
    /// Appending to a document that already has a bullet list should reuse its numbering rather than
    /// adding a near-duplicate definition.
    /// </summary>
    [Fact]
    public void AddListItem_AfterReopening_ReusesTheExistingNumbering()
    {
        var filePath = GetFilepath("list-append.docx");
        using (var document = CreateDocument(filePath))
        {
            document.GetBody().AddListItem("First");
        }

        using (var document = OpenDocument(filePath))
        {
            document.GetBody().AddListItem("Second");
        }

        OpenXmlValidation.AssertValid(filePath);

        using var package = WordprocessingDocument.Open(filePath, false);
        var numbering = package.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;

        Assert.Single(numbering.Elements<WordLib.AbstractNum>());
        Assert.Single(numbering.Elements<WordLib.NumberingInstance>());
    }

    /// <summary>
    /// Numbering id 0 is the format's reserved "not in a list" value, so removing a paragraph from a
    /// list is an explicit statement rather than the absence of one.
    /// </summary>
    [Fact]
    public void ListStyleNone_RemovesTheParagraphFromItsList()
    {
        using var stream = WriteAndValidate(body =>
        {
            var paragraph = body.AddListItem("Was a list item");
            paragraph.ApplyFormat(new ParagraphFormat { ListStyle = ListStyle.None });
        });

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal(ListStyle.None, document.GetBody().Paragraphs[0].Format.ListStyle);
    }

    [Fact]
    public void AddParagraph_WithoutAList_CreatesNoNumberingPart()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph("plain"));

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);

        Assert.Null(package.MainDocumentPart!.NumberingDefinitionsPart);
    }

    [Fact]
    public void AddListItem_InATableCell_IsSchemaValid()
    {
        using var stream = WriteAndValidate(body =>
        {
            var cell = body.AddTable(1, 1).GetCell(0, 0);
            cell.AddListItem("inside a cell");
        });

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal("inside a cell", document.GetBody().Tables[0].GetAllTexts());
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(9)]
    public void ListLevel_OutOfRange_ThrowsOnAssignment(int level)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ParagraphFormat { ListLevel = level });
    }

    [Fact]
    public void AddListItem_WithLevelOutOfRange_Throws()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        Assert.Throws<ArgumentOutOfRangeException>(() => body.AddListItem("x", ListStyle.Bullet, level: 9));
    }
}
