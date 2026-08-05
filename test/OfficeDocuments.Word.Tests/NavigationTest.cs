using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.TestKit;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Walking an existing document and editing its structure.
/// </summary>
public class NavigationTest : WordTestBase
{
    /// <summary>
    /// Document order, not paragraphs-then-tables. A paragraph written between two tables has to be
    /// reported between them, because the order is what a caller reading the document relies on.
    /// </summary>
    [Fact]
    public void GetAllParagraphs_ReturnsDocumentOrderAndDescendsIntoTables()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        body.AddParagraph("before");
        body.AddTable([["cell one", "cell two"]]);
        body.AddParagraph("between");
        body.AddTable([["cell three"]]);
        body.AddParagraph("after");

        var texts = body.GetAllParagraphs().Select(paragraph => paragraph.GetTexts()).ToList();

        Assert.Equal(["before", "cell one", "cell two", "between", "cell three", "after"], texts);
    }

    [Fact]
    public void GetAllParagraphs_DescendsIntoNestedTables()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        var outer = body.AddTable(1, 1);
        outer.GetCell(0, 0).SetText("outer");
        outer.GetCell(0, 0).AddTable([["inner"]]);

        var texts = body.GetAllParagraphs().Select(paragraph => paragraph.GetTexts()).ToList();

        Assert.Equal(["outer", "inner"], texts);
    }

    /// <summary>
    /// <c>Paragraphs</c> stops at this container's own children; <c>GetAllParagraphs</c> does not. The
    /// distinction is what lets a caller choose between editing the body's own text and sweeping
    /// everything.
    /// </summary>
    [Fact]
    public void Paragraphs_AndGetAllParagraphs_DifferOnTableContent()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        body.AddParagraph("body text");
        body.AddTable([["in a cell"]]);

        Assert.Single(body.Paragraphs);
        Assert.Equal(2, body.GetAllParagraphs().Count());
    }

    [Fact]
    public void FindParagraphs_ReturnsTheMatchesIncludingInsideTables()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        body.AddParagraph("invoice 2026-07");
        body.AddParagraph("unrelated");
        body.AddTable([["invoice 2026-08"]]);

        var found = body.FindParagraphs("invoice").Select(paragraph => paragraph.GetTexts()).ToList();

        Assert.Equal(["invoice 2026-07", "invoice 2026-08"], found);
    }

    [Fact]
    public void FindParagraphs_IgnoringCase_Matches()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();
        body.AddParagraph("Summary");

        Assert.Single(body.FindParagraphs("summary", StringComparison.OrdinalIgnoreCase));
        Assert.Empty(body.FindParagraphs("summary"));
    }

    /// <summary>
    /// The regression this whole collection model exists for: a removal has to be visible in the
    /// projected list immediately, without anyone having to keep a second copy in step.
    /// </summary>
    [Fact]
    public void Remove_Paragraph_IsVisibleInTheProjectedListAtOnce()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        body.AddParagraph("keep");
        var doomed = body.AddParagraph("remove");
        body.AddParagraph("keep too");

        Assert.Equal(3, body.Paragraphs.Count);
        Assert.True(body.Remove(doomed));

        Assert.Equal(2, body.Paragraphs.Count);
        Assert.Equal("keep\nkeep too", body.GetAllTexts());
    }

    [Fact]
    public void Remove_Table_TakesItsContentWithIt()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            var body = document.GetBody();
            body.AddParagraph("kept");
            var table = body.AddTable([["gone"]]);

            Assert.True(body.Remove(table));
            Assert.Empty(body.Tables);
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);
        Assert.Equal("kept", reopened.GetBody().GetAllTexts());
    }

    /// <summary>
    /// Removing something that belongs to a different container is refused rather than performed. It
    /// would otherwise be an edit to a part of the document the caller did not name.
    /// </summary>
    [Fact]
    public void Remove_SomethingFromAnotherContainer_ReturnsFalseAndChangesNothing()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();
        var cell = body.AddTable(1, 1).GetCell(0, 0);
        var cellParagraph = cell.AddParagraph("in the cell");

        Assert.False(body.Remove(cellParagraph));
        Assert.Equal(2, cell.Paragraphs.Count);
    }

    [Fact]
    public void Remove_TableRow_LeavesTheRestOfTheTableValid()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            var table = document.GetBody().AddTable([
                ["Header", "Value"],
                ["draft", "0"],
                ["final", "1"],
            ]);

            Assert.True(table.Remove(table.Rows[1]));
            Assert.Equal(2, table.Rows.Count);
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);
        var read = reopened.GetBody().Tables[0];

        Assert.Equal(2, read.Rows.Count);
        Assert.Equal("Header\tValue\nfinal\t1", read.GetAllTexts());
    }

    [Fact]
    public void Remove_RowFromAnotherTable_ReturnsFalse()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();
        var first = body.AddTable([["a"]]);
        var second = body.AddTable([["b"]]);

        Assert.False(first.Remove(second.Rows[0]));
        Assert.Single(second.Rows);
    }

    /// <summary>
    /// The bug that motivated dropping the cached collection: setting a cell's text removes its
    /// paragraph and adds another, so a list built before that call reported both.
    /// </summary>
    [Fact]
    public void SetText_OnACellWhoseParagraphsWereAlreadyRead_DoesNotDoubleCount()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var cell = document.GetBody().AddTable(1, 1).GetCell(0, 0);

        Assert.Single(cell.Paragraphs);

        cell.SetText("replaced");

        Assert.Single(cell.Paragraphs);
        Assert.Equal("replaced", cell.GetAllTexts());
    }

    /// <summary>
    /// A caller holding a reference to a paragraph keeps getting the same instance back, so identity
    /// comparisons against the projected list work.
    /// </summary>
    [Fact]
    public void Paragraphs_ReturnTheSameInstanceAcrossReads()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();
        var added = body.AddParagraph("text");

        Assert.Same(added, body.Paragraphs[0]);
        Assert.Same(body.Paragraphs[0], body.Paragraphs[0]);
        Assert.Same(added, body.GetAllParagraphs().Single());
    }

    /// <summary>
    /// Reading a foreign document has to work through the same projections: the paragraphs, the runs
    /// Word split, and the section properties it always writes.
    /// </summary>
    [Fact]
    public void OpenForeignDocument_ProjectsItsParagraphsAndRuns()
    {
        using var input = ForeignDocuments.WithSplitRuns(
            ["A single run paragraph"],
            ["Split ", "across ", "three runs"]);

        using var document = OpenDocument(input, isEditable: false);
        var paragraphs = document.GetBody().Paragraphs;

        Assert.Equal(2, paragraphs.Count);
        Assert.Single(paragraphs[0].Runs);
        Assert.Equal(3, paragraphs[1].Runs.Count);
        Assert.Equal("Split across three runs", paragraphs[1].GetTexts());
        Assert.Equal("A single run paragraph\nSplit across three runs", document.GetBody().GetAllTexts());
    }

    /// <summary>
    /// Appending to a document that already has section properties must not put content after them.
    /// The document is valid either way as far as a round trip through this library can tell; only the
    /// schema — and Word — reject it.
    /// </summary>
    [Fact]
    public void AppendToForeignDocument_KeepsSectionPropertiesLast()
    {
        using var input = ForeignDocuments.WithSplitRuns(["existing"]);

        using (var document = OpenDocument(input))
        {
            document.GetBody().AddParagraph("appended", new ParagraphFormat { StyleId = WordStyleIds.Heading1 });
            document.GetBody().AddTable([["also appended"]]);
        }

        OpenXmlValidation.AssertValid(input);

        var lastChildName = ReadDocumentElement(input, element => element.Body!.LastChild?.LocalName);

        Assert.Equal("sectPr", lastChildName);
    }
}
