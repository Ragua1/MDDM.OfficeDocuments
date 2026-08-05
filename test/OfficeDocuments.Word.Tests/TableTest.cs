using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Covers table creation, filling, formatting, and reading.
/// </summary>
public class TableTest : WordTestBase
{
    [Fact]
    public void AddTable_WithSize_CreatesEveryRowAndCell()
    {
        using var stream = WriteAndValidate(body => body.AddTable(3, 4));

        using var document = OpenDocument(stream, isEditable: false);
        var table = document.GetBody().Tables[0];

        Assert.Equal(4, table.ColumnCount);
        Assert.Equal(3, table.Rows.Count);
        Assert.All(table.Rows, row => Assert.Equal(4, row.Cells.Count));
    }

    /// <summary>
    /// <c>CT_Tbl</c> requires a <c>w:tblGrid</c> declaring one column per grid column; without it Word
    /// cannot lay the table out.
    /// </summary>
    [Fact]
    public void AddTable_DeclaresOneGridColumnPerColumn()
    {
        using var stream = WriteAndValidate(body => body.AddTable(1, 3));

        var grid = ReadDocumentElement(stream, document => document.Descendants<WordLib.TableGrid>().Single());

        Assert.Equal(3, grid.Elements<WordLib.GridColumn>().Count());
    }

    /// <summary>
    /// <c>CT_Tc</c> requires block content, so a cell with no paragraph makes Word offer to repair the
    /// document. The schema gate catches it, and this test says why it matters.
    /// </summary>
    [Fact]
    public void AddTable_GivesEveryCellAParagraph()
    {
        using var stream = WriteAndValidate(body => body.AddTable(2, 2));

        var cells = ReadDocumentElement(stream, document => document.Descendants<WordLib.TableCell>().ToList());

        Assert.Equal(4, cells.Count);
        Assert.All(cells, cell => Assert.NotEmpty(cell.Elements<WordLib.Paragraph>()));
    }

    [Fact]
    public void AddTable_FromData_FillsCellsInOrder()
    {
        string[][] rows =
        [
            ["Item", "Quantity", "Price"],
            ["Widget", "2", "19.90"],
            ["Gadget", "1", "45.00"],
        ];

        using var stream = WriteAndValidate(body => body.AddTable(rows));

        using var document = OpenDocument(stream, isEditable: false);
        var table = document.GetBody().Tables[0];

        Assert.Equal(3, table.ColumnCount);
        Assert.Equal("Item\tQuantity\tPrice", table.Rows[0].GetAllTexts());
        Assert.Equal("Gadget\t1\t45.00", table.Rows[2].GetAllTexts());
        Assert.Equal("Widget", table.GetCell(1, 0).GetAllTexts());
    }

    /// <summary>
    /// A ragged input would otherwise produce rows narrower than the grid, which Word renders as a
    /// broken table.
    /// </summary>
    [Fact]
    public void AddTable_FromRaggedData_PadsEveryRowToTheGridWidth()
    {
        string[][] rows =
        [
            ["a", "b", "c"],
            ["d"],
        ];

        using var stream = WriteAndValidate(body => body.AddTable(rows));

        using var document = OpenDocument(stream, isEditable: false);
        var table = document.GetBody().Tables[0];

        Assert.Equal(3, table.ColumnCount);
        Assert.All(table.Rows, row => Assert.Equal(3, row.Cells.Count));
        Assert.Equal("d\t\t", table.Rows[1].GetAllTexts());
    }

    [Fact]
    public void ApplyFormat_RoundTripsEveryModelledProperty()
    {
        var format = new TableFormat
        {
            WidthPercent = 80,
            Alignment = TableAlignment.Center,
            Borders = TableBorders.All,
            BorderColor = "#4472c4",
            BorderWidth = 1.5,
            CellPadding = 4,
        };

        using var stream = WriteAndValidate(body => body.AddTable(1, 2, format));

        using var document = OpenDocument(stream, isEditable: false);
        var actual = document.GetBody().Tables[0].Format;

        Assert.Equal(80, actual.WidthPercent);
        Assert.Equal(TableAlignment.Center, actual.Alignment);
        Assert.Equal(TableBorders.All, actual.Borders);
        Assert.Equal("4472C4", actual.BorderColor);
        Assert.Equal(1.5, actual.BorderWidth);
        Assert.Equal(4, actual.CellPadding);
    }

    /// <summary>
    /// Outline borders have to write the inside borders as an explicit "none", otherwise a table style
    /// can put the grid lines back.
    /// </summary>
    [Fact]
    public void ApplyFormat_WithOutlineBorders_SuppressesInsideBorders()
    {
        using var stream = WriteAndValidate(body => body.AddTable(2, 2, new TableFormat { Borders = TableBorders.Outline }));

        var borders = ReadDocumentElement(stream, document => document.Descendants<WordLib.TableBorders>().First());

        Assert.Equal(WordLib.BorderValues.Single, borders.TopBorder?.Val?.Value);
        Assert.Equal(WordLib.BorderValues.None, borders.InsideHorizontalBorder?.Val?.Value);
        Assert.Equal(WordLib.BorderValues.None, borders.InsideVerticalBorder?.Val?.Value);
    }

    [Fact]
    public void ApplyFormat_WithNoBorders_RoundTripsAsNone()
    {
        using var stream = WriteAndValidate(body => body.AddTable(1, 1, new TableFormat { Borders = TableBorders.None }));

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal(TableBorders.None, document.GetBody().Tables[0].Format.Borders);
    }

    [Fact]
    public void CellFormat_RoundTripsEveryModelledProperty()
    {
        var format = new TableCellFormat
        {
            WidthPercent = 25,
            BackgroundColor = "D9E2F3",
            VerticalAlignment = CellVerticalAlignment.Center,
            ColumnSpan = 2,
        };

        using var stream = WriteAndValidate(body =>
        {
            var table = body.AddTable(1, 3);
            table.GetCell(0, 0).ApplyFormat(format);
        });

        using var document = OpenDocument(stream, isEditable: false);
        var actual = document.GetBody().Tables[0].GetCell(0, 0).Format;

        Assert.Equal(25, actual.WidthPercent);
        Assert.Equal("D9E2F3", actual.BackgroundColor);
        Assert.Equal(CellVerticalAlignment.Center, actual.VerticalAlignment);
        Assert.Equal(2, actual.ColumnSpan);
    }

    /// <summary>
    /// Shading needs an explicit pattern; a fill colour with no <c>w:val</c> has no visible effect.
    /// </summary>
    [Fact]
    public void CellBackground_WritesAClearShadingPattern()
    {
        using var stream = WriteAndValidate(body =>
            body.AddTable(1, 1).GetCell(0, 0).ApplyFormat(new TableCellFormat { BackgroundColor = "FF0000" }));

        var shading = ReadDocumentElement(stream, document => document.Descendants<WordLib.Shading>().Single());

        Assert.Equal(WordLib.ShadingPatternValues.Clear, shading.Val?.Value);
        Assert.Equal("FF0000", shading.Fill?.Value);
    }

    [Fact]
    public void RepeatAsHeader_MarksTheRowAndRoundTrips()
    {
        using var stream = WriteAndValidate(body =>
        {
            var table = body.AddTable([["Header"], ["Body"]]);
            table.Rows[0].RepeatAsHeader();
        });

        using var document = OpenDocument(stream, isEditable: false);
        var table = document.GetBody().Tables[0];

        Assert.True(table.Rows[0].IsHeader);
        Assert.False(table.Rows[1].IsHeader);
    }

    /// <summary>
    /// A cell is a block container, so everything the body can hold works inside one.
    /// </summary>
    [Fact]
    public void TableCell_AcceptsFullBlockContent()
    {
        using var stream = WriteAndValidate(body =>
        {
            var cell = body.AddTable(1, 1).GetCell(0, 0);
            cell.SetText("first");
            cell.AddParagraph("second", new ParagraphFormat { Alignment = ParagraphAlignment.Right });
            cell.AddListItem("bullet");
        });

        using var document = OpenDocument(stream, isEditable: false);
        var cell = document.GetBody().Tables[0].GetCell(0, 0);

        Assert.Equal(3, cell.Paragraphs.Count);
        Assert.Equal("first\nsecond\nbullet", cell.GetAllTexts());
    }

    [Fact]
    public void TableCell_NestedTable_IsSchemaValid()
    {
        using var stream = WriteAndValidate(body =>
        {
            var cell = body.AddTable(1, 1).GetCell(0, 0);
            cell.AddTable([["nested"]]);
        });

        using var document = OpenDocument(stream, isEditable: false);
        var outerCell = document.GetBody().Tables[0].GetCell(0, 0);

        Assert.Single(outerCell.Tables);
        Assert.Equal("nested", outerCell.Tables[0].GetAllTexts());
    }

    /// <summary>
    /// SetText replaces content rather than appending, but a cell must never end up with no block
    /// content at all.
    /// </summary>
    [Fact]
    public void SetText_ReplacesExistingContentAndKeepsOneParagraph()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var cell = document.GetBody().AddTable(1, 1).GetCell(0, 0);
        cell.AddParagraph("extra");

        cell.SetText("only");

        Assert.Single(cell.Paragraphs);
        Assert.Equal("only", cell.GetAllTexts());
    }

    /// <summary>
    /// The body reads its blocks in document order, so a table between two paragraphs appears between
    /// them rather than after everything else.
    /// </summary>
    [Fact]
    public void GetAllTexts_ReadsParagraphsAndTablesInDocumentOrder()
    {
        using var stream = WriteAndValidate(body =>
        {
            body.AddParagraph("before");
            body.AddTable([["a", "b"]]);
            body.AddParagraph("after");
        });

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal("before\na\tb\nafter", document.GetBody().GetAllTexts());
    }

    [Fact]
    public void AddTable_AfterOpeningADocument_KeepsSectionPropertiesLast()
    {
        var filePath = GetFilepath("table-append.docx");
        using (var document = CreateDocument(filePath))
        {
            document.GetBody().AddParagraph("existing");
        }

        using (var document = OpenDocument(filePath))
        {
            document.GetBody().AddTable([["appended"]]);
        }

        OpenXmlValidation.AssertValid(filePath);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    [InlineData(-1, 2)]
    public void AddTable_WithInvalidSize_Throws(int rowCount, int columnCount)
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        Assert.Throws<ArgumentOutOfRangeException>(() => body.AddTable(rowCount, columnCount));
    }

    [Fact]
    public void AddTable_WithNoRows_Throws()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        Assert.Throws<ArgumentException>(() => body.AddTable([]));
    }

    [Fact]
    public void GetCell_OutOfRange_Throws()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var table = document.GetBody().AddTable(1, 1);

        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetCell(1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetCell(0, 1));
    }

    [Theory]
    [InlineData(-1d)]
    [InlineData(101d)]
    public void TableWidthPercent_OutOfRange_ThrowsOnAssignment(double widthPercent)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new TableFormat { WidthPercent = widthPercent });
    }

    [Fact]
    public void ColumnSpan_BelowOne_ThrowsOnAssignment()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new TableCellFormat { ColumnSpan = 0 });
    }
}
