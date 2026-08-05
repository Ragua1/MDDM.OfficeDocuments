using OfficeDocuments.Excel.TestKit;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Options;
using OfficeDocuments.Excel.Styles;
using XdrSpr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Color = System.Drawing.Color;
using StyleFill = OfficeDocuments.Excel.Styles.Fill;

namespace OfficeDocuments.Excel.IntegrationTests;

public class RangeAndAdvancedFeaturesTest : SpreadsheetTestBase
{
    [Fact]
    public void GetRange_SetValuesAndMerge_WritesExpectedCells()
    {
        using var stream = new MemoryStream();
        using (var spreadsheet = CreateNewSpreadsheet(stream))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet 1");
            var range = worksheet.GetRange("B2:C3");

            range.SetValues(
            [
                [1, 2],
                [3, 4]
            ]);
            range.Merge();

            var values = range.GetValues();

            Assert.Equal("1", values[0][0]);
            Assert.Equal("4", values[1][1]);
        }

        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = WorkbookParts.GetWorksheetPart(document, "Sheet 1");
        var worksheetElement = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet element was not found.");
        var mergeCells = worksheetElement.GetFirstChild<MergeCells>() ?? throw new InvalidOperationException("MergeCells element was not found.");
        var mergeCell = mergeCells.Elements<MergeCell>().SingleOrDefault();
        Assert.NotNull(mergeCell);
        Assert.Equal("B2:C3", mergeCell.Reference?.Value);
    }

    [Fact]
    public void GetRange_Worksheet_IsAccessibleThroughInterface()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        IWorksheet worksheet = spreadsheet.AddWorksheet("Sheet 1");

        IRange range = worksheet.GetRange("B2:C3");

        Assert.NotNull(range.Worksheet);
        Assert.Same(worksheet, range.Worksheet);
    }

    [Fact]
    public void SortByColumn_WithHeader_SortsBodyOnly()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet 1");
        var range = worksheet.GetRange("A1:B4");

        range.SetValues(
        [
            ["Name", "Score"],
            ["Alice", 10],
            ["Bob", 30],
            ["Cara", 20]
        ]);

        range.SortByColumn(2, SortDirection.Descending, hasHeader: true);

        Assert.Equal("Name", worksheet.GetCell(1, 1)?.Value);
        Assert.Equal("Bob", worksheet.GetCell(1, 2)?.Value);
        Assert.Equal("Cara", worksheet.GetCell(1, 3)?.Value);
        Assert.Equal("Alice", worksheet.GetCell(1, 4)?.Value);
    }

    [Fact]
    public void AddRows_WithObjectsAndHeader_WritesExpectedRange()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet 1");

        var range = worksheet.AddRows(
            [
                new ExportRow("Alice", 30),
                new ExportRow("Bob", 25)
            ],
            includeHeader: true);

        Assert.NotNull(range);
        Assert.Equal("Name", worksheet.GetCell(1, 1)?.Value);
        Assert.Equal("Age", worksheet.GetCell(2, 1)?.Value);
        Assert.Equal("Alice", worksheet.GetCell(1, 2)?.Value);
        Assert.Equal("30", worksheet.GetCell(2, 2)?.Value);
        Assert.Equal("Bob", worksheet.GetCell(1, 3)?.Value);
    }

    [Fact]
    public void WorksheetOperations_RenameMoveCopyHideAndRemove_WorkAsExpected()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var first = spreadsheet.AddWorksheet("First");
        first.AddCell(1, 1, "Header");
        spreadsheet.AddWorksheet("Second");
        spreadsheet.AddWorksheet("Third");

        spreadsheet.RenameWorksheet("Second", "Renamed");
        spreadsheet.MoveWorksheet("Renamed", 1);
        var clone = spreadsheet.CopyWorksheet("Renamed", "Renamed Copy");
        spreadsheet.SetWorksheetHidden("Renamed Copy", true);
        spreadsheet.RemoveWorksheet("Third");

        Assert.Equal(["Renamed", "Renamed Copy", "First"], spreadsheet.GetWorksheetsName().ToArray());
        Assert.True(clone.IsHidden);
    }

    [Fact]
    public void WorksheetUsabilityFeatures_FreezeAndAutoFit_UpdateWorksheetMetadata()
    {
        using var stream = new MemoryStream();
        using (var spreadsheet = CreateNewSpreadsheet(stream))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet 1");
            worksheet.AddRows(
            [
                ["Very long heading", "X"],
                ["Short", "Y"]
            ]);

            worksheet.FreezePanes(1, 1);
            worksheet.AutoFitColumns();
            worksheet.ClearFrozenPanes();
        }

        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = WorkbookParts.GetWorksheetPart(document, "Sheet 1");
        var worksheetElement = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet element was not found.");
        var columns = worksheetElement.GetFirstChild<Columns>() ?? throw new InvalidOperationException("Columns element was not found.");
        Assert.Contains(columns.Elements<Column>(), column => column.Width is { Value: > 0 });
        Assert.Null(worksheetElement.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.GetFirstChild<Pane>());
    }

    [Fact]
    public void ValidationAndConditionalFormatting_CreateExpectedWorksheetNodes()
    {
        using var stream = new MemoryStream();
        using (var spreadsheet = CreateNewSpreadsheet(stream))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet 1");
            var range = worksheet.GetRange("A1:A3");
            var style = spreadsheet.CreateStyle(fill: new StyleFill(Color.LightGoldenrodYellow));

            range.SetValues(
            [
                ["A"],
                ["B"],
                ["C"]
            ]);
            range.AddValidation(DataValidationOptions.List(["A", "B", "C"]));
            range.AddConditionalFormatting(ConditionalFormattingOptions.ContainsText("A", style));
        }

        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = WorkbookParts.GetWorksheetPart(document, "Sheet 1");
        var worksheetElement = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet element was not found.");
        var dataValidations = worksheetElement.GetFirstChild<DataValidations>() ?? throw new InvalidOperationException("DataValidations element was not found.");
        Assert.NotNull(dataValidations);
        Assert.NotEmpty(worksheetElement.Elements<ConditionalFormatting>());
    }


    [Fact]
    public void AddTable_WithValidRange_ReturnsTableInfo()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        worksheet.AddRows([["Name", "Age"], ["Alice", 30], ["Bob", 25]]);
        var start = worksheet.AddCellOnIndex(1, 1);
        var end = worksheet.AddCellOnIndex(2, 3);

        var info = spreadsheet.AddTable("Sheet1", start, end, ["Name", "Age"]);

        Assert.NotNull(info);
        Assert.Equal("Name", info.ColumnNames[0]);
        Assert.Equal("Age", info.ColumnNames[1]);
        Assert.Equal(2, info.ColumnCount);
        Assert.Equal("Sheet1", info.WorksheetName);
        Assert.Contains("A1", info.Reference);
    }

    [Fact]
    public void AddTable_WithOptions_UsesProvidedNameAndStyle()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        worksheet.AddRows([["Product", "Price"], ["Widget", 9.99]]);
        var range = worksheet.GetRange("A1:B2");
        var opts = new TableCreateOptions
        {
            TableName = "SalesTable",
            DisplayName = "SalesTable",
            Style = new TableStyleOptions { StyleName = "TableStyleLight1", ShowBandedRows = true }
        };

        var info = spreadsheet.AddTable(range, ["Product", "Price"], opts);

        Assert.Equal("SalesTable", info.Name);
        Assert.Equal("SalesTable", info.DisplayName);
    }

    [Fact]
    public void AddTable_ColumnCountMismatch_ThrowsArgumentException()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        worksheet.AddRows([["A", "B", "C"], ["1", "2", "3"]]);
        var start = worksheet.AddCellOnIndex(1, 1);
        var end = worksheet.AddCellOnIndex(3, 2);

        // Provide only 2 column names for 3-column range
        Assert.Throws<ArgumentException>(() =>
            spreadsheet.AddTable("Sheet1", start, end, ["Only", "Two"]));
    }

    [Fact]
    public void GetTable_ByName_ReturnsExpectedTableInfo()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        worksheet.AddRows([["X", "Y"], ["1", "2"]]);
        var start = worksheet.AddCellOnIndex(1, 1);
        var end = worksheet.AddCellOnIndex(2, 2);
        spreadsheet.AddTable("Sheet1", start, end, ["X", "Y"]);

        var info = spreadsheet.GetTable("Sheet1", "Table1");

        Assert.NotNull(info);
        Assert.Equal("Table1", info.Name);
        Assert.Equal(2, info.ColumnCount);
    }

    [Fact]
    public void GetTable_NonExistentName_ReturnsNull()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        spreadsheet.AddWorksheet("Sheet1");

        var info = spreadsheet.GetTable("Sheet1", "NoSuchTable");

        Assert.Null(info);
    }

    [Fact]
    public void GetTables_WithWorksheetName_ReturnsAllTablesOnSheet()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        worksheet.AddRows([["A", "B"], ["1", "2"], ["3", "4"], ["5", "6"]]);
        var s1 = worksheet.AddCellOnIndex(1, 1);
        var e1 = worksheet.AddCellOnIndex(2, 2);
        var s2 = worksheet.AddCellOnIndex(1, 3);
        var e2 = worksheet.AddCellOnIndex(2, 4);
        spreadsheet.AddTable("Sheet1", s1, e1, ["A", "B"]);
        spreadsheet.AddTable("Sheet1", s2, e2, ["A", "B"]);

        var tables = spreadsheet.GetTables("Sheet1").ToList();

        Assert.Equal(2, tables.Count);
    }

    [Fact]
    public void GetTables_AllWorksheets_ReturnsTablesFromAllSheets()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var ws1 = spreadsheet.AddWorksheet("Sheet1");
        var ws2 = spreadsheet.AddWorksheet("Sheet2");
        ws1.AddRows([["A", "B"], ["1", "2"]]);
        ws2.AddRows([["C", "D"], ["3", "4"]]);
        spreadsheet.AddTable("Sheet1", ws1.AddCellOnIndex(1, 1), ws1.AddCellOnIndex(2, 2), ["A", "B"]);
        spreadsheet.AddTable("Sheet2", ws2.AddCellOnIndex(1, 1), ws2.AddCellOnIndex(2, 2), ["C", "D"]);

        var all = spreadsheet.GetTables().ToList();

        Assert.Equal(2, all.Count);
        Assert.Contains(all, t => t.WorksheetName == "Sheet1");
        Assert.Contains(all, t => t.WorksheetName == "Sheet2");
    }

    [Fact]
    public void RenameTable_ChangesTableName()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        worksheet.AddRows([["Col1", "Col2"], ["v1", "v2"]]);
        spreadsheet.AddTable("Sheet1", worksheet.AddCellOnIndex(1, 1), worksheet.AddCellOnIndex(2, 2), ["Col1", "Col2"]);

        spreadsheet.RenameTable("Sheet1", "Table1", "RenamedTable");

        Assert.Null(spreadsheet.GetTable("Sheet1", "Table1"));
        var renamed = spreadsheet.GetTable("Sheet1", "RenamedTable");
        Assert.NotNull(renamed);
        Assert.Equal("RenamedTable", renamed.Name);
    }

    [Fact]
    public void RenameTable_DuplicateName_ThrowsArgumentException()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        worksheet.AddRows([["A", "B"], ["1", "2"], ["3", "4"], ["5", "6"]]);
        var s1 = worksheet.AddCellOnIndex(1, 1);
        var e1 = worksheet.AddCellOnIndex(2, 2);
        var s2 = worksheet.AddCellOnIndex(1, 3);
        var e2 = worksheet.AddCellOnIndex(2, 4);
        spreadsheet.AddTable("Sheet1", s1, e1, ["A", "B"]);
        spreadsheet.AddTable("Sheet1", s2, e2, ["A", "B"]);

        // Table1 exists, Table2 exists - try to rename Table1 to Table2
        Assert.Throws<ArgumentException>(() =>
            spreadsheet.RenameTable("Sheet1", "Table1", "Table2"));
    }

    [Fact]
    public void ResizeTable_WithSameColumnCount_UpdatesReference()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        worksheet.AddRows([["A", "B"], ["1", "2"], ["3", "4"], ["5", "6"]]);
        spreadsheet.AddTable("Sheet1", worksheet.AddCellOnIndex(1, 1), worksheet.AddCellOnIndex(2, 2), ["A", "B"]);
        var newRange = worksheet.GetRange("A1:B4");

        spreadsheet.ResizeTable("Sheet1", "Table1", newRange);

        var info = spreadsheet.GetTable("Sheet1", "Table1");
        Assert.NotNull(info);
        Assert.Contains("B4", info.Reference);
    }

    [Fact]
    public void ResizeTable_WithDifferentColumnCount_ThrowsArgumentException()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        worksheet.AddRows([["A", "B"], ["1", "2"], ["3", "4"]]);
        spreadsheet.AddTable("Sheet1", worksheet.AddCellOnIndex(1, 1), worksheet.AddCellOnIndex(2, 2), ["A", "B"]);
        var wrongRange = worksheet.GetRange("A1:C3"); // 3 columns, table has 2

        Assert.Throws<ArgumentException>(() =>
            spreadsheet.ResizeTable("Sheet1", "Table1", wrongRange));
    }

    [Fact]
    public void RemoveTable_RemovesTableFromWorksheet()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        worksheet.AddRows([["A", "B"], ["1", "2"]]);
        spreadsheet.AddTable("Sheet1", worksheet.AddCellOnIndex(1, 1), worksheet.AddCellOnIndex(2, 2), ["A", "B"]);

        spreadsheet.RemoveTable("Sheet1", "Table1");

        Assert.Null(spreadsheet.GetTable("Sheet1", "Table1"));
        Assert.Empty(spreadsheet.GetTables("Sheet1"));
    }

    [Fact]
    public void RemoveTable_NonExistentTable_ThrowsArgumentException()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        spreadsheet.AddWorksheet("Sheet1");

        Assert.Throws<ArgumentException>(() =>
            spreadsheet.RemoveTable("Sheet1", "NoSuchTable"));
    }

    [Fact]
    public void AddTable_ViaRange_WithOptions_ReturnsTableInfoWithCorrectWorksheet()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        worksheet.AddRows([["Item", "Qty"], ["Apple", 5]]);
        var range = worksheet.GetRange("A1:B2");

        var info = spreadsheet.AddTable(range, ["Item", "Qty"], new TableCreateOptions { TableName = "FruitTable" });

        Assert.Equal("FruitTable", info.Name);
        Assert.Equal("Sheet1", info.WorksheetName);
        Assert.Equal(2, info.ColumnCount);
    }

    private sealed record ExportRow(string Name, int Age);

    #region AddImage tests


    [Fact]
    public void AddImage_FromStream_CreatesDrawingsPartAndAnchor()
    {
        using var stream = new MemoryStream();
        using (var spreadsheet = CreateNewSpreadsheet(stream))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet1");
            using var imageStream = new MemoryStream(TestImages.MinimalPng());

            worksheet.AddImage(imageStream, ImageType.Png, 2, 3, 4, 6);
        }

        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = WorkbookParts.GetWorksheetPart(document, "Sheet1");
        var worksheetElement = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet element was not found.");
        var drawing = worksheetElement.GetFirstChild<Drawing>() ?? throw new InvalidOperationException("Drawing element was not found.");
        Assert.NotNull(drawing);

        var drawingsPart = worksheetPart.DrawingsPart;
        Assert.NotNull(drawingsPart);
        var worksheetDrawing = drawingsPart.WorksheetDrawing;
        Assert.NotNull(worksheetDrawing);
        var anchor = worksheetDrawing.Elements<XdrSpr.TwoCellAnchor>().SingleOrDefault();
        Assert.NotNull(anchor);
        var from = anchor.GetFirstChild<XdrSpr.FromMarker>();
        Assert.NotNull(from);
        var columnId = from.GetFirstChild<XdrSpr.ColumnId>();
        var rowId = from.GetFirstChild<XdrSpr.RowId>();
        Assert.NotNull(columnId);
        Assert.NotNull(rowId);
        // fromColumn=2 → 0-based index = 1
        Assert.Equal("1", columnId.Text);
        // fromRow=3 → 0-based index = 2
        Assert.Equal("2", rowId.Text);
    }

    [Fact]
    public void AddImage_FromFilePath_AutoDetectsTypeAndEmbeds()
    {
        using var stream = new MemoryStream();
        var imgFile = GetFilepath("test-image.png");
        File.WriteAllBytes(imgFile, TestImages.MinimalPng());

        using (var spreadsheet = CreateNewSpreadsheet(stream))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet1");
            worksheet.AddImage(imgFile, 1, 1, 3, 3);
        }

        using var document = SpreadsheetDocument.Open(stream, false);
        var drawingsPart = WorkbookParts.GetWorksheetPart(document, "Sheet1").DrawingsPart;
        Assert.NotNull(drawingsPart);
        Assert.Single(drawingsPart.ImageParts);
    }

    [Fact]
    public void AddImage_MultipleImages_AllAnchorsPresent()
    {
        using var stream = new MemoryStream();
        using (var spreadsheet = CreateNewSpreadsheet(stream))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet1");
            var pngBytes = TestImages.MinimalPng();

            using var s1 = new MemoryStream(pngBytes);
            using var s2 = new MemoryStream(pngBytes);
            using var s3 = new MemoryStream(pngBytes);
            worksheet.AddImage(s1, ImageType.Png, 1, 1, 2, 2);
            worksheet.AddImage(s2, ImageType.Png, 3, 1, 4, 2);
            worksheet.AddImage(s3, ImageType.Png, 1, 3, 2, 4);
        }

        using var document = SpreadsheetDocument.Open(stream, false);
        var anchors = WorkbookParts.GetWorksheetPart(document, "Sheet1").DrawingsPart?.WorksheetDrawing?
            .Elements<XdrSpr.TwoCellAnchor>().ToList();
        Assert.NotNull(anchors);
        Assert.Equal(3, anchors.Count);
    }


    [Fact]
    public void AddImage_InvalidFromColumn_ThrowsArgumentException()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        using var stream = new MemoryStream(TestImages.MinimalPng());

        Assert.Throws<ArgumentException>(() => worksheet.AddImage(stream, ImageType.Png, 0, 1, 2, 3));
    }

    [Fact]
    public void AddImage_ToColumnLessThanFromColumn_ThrowsArgumentException()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        using var stream = new MemoryStream(TestImages.MinimalPng());

        Assert.Throws<ArgumentException>(() => worksheet.AddImage(stream, ImageType.Png, 3, 1, 2, 3));
    }

    [Fact]
    public void AddImage_ToRowLessThanFromRow_ThrowsArgumentException()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        using var stream = new MemoryStream(TestImages.MinimalPng());

        Assert.Throws<ArgumentException>(() => worksheet.AddImage(stream, ImageType.Png, 1, 3, 2, 2));
    }

    [Fact]
    public void AddImage_UnsupportedExtension_ThrowsArgumentException()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        var imgFile = GetFilepath("test-image.svg");
        File.WriteAllBytes(imgFile, "<svg/>"u8.ToArray());

        Assert.Throws<ArgumentException>(() => worksheet.AddImage(imgFile, 1, 1, 3, 3));
    }

    [Fact]
    public void AddImage_FileNotFound_ThrowsFileNotFoundException()
    {
        using var spreadsheet = CreateInMemorySpreadsheet();
        var worksheet = spreadsheet.AddWorksheet("Sheet1");

        Assert.Throws<FileNotFoundException>(() => worksheet.AddImage("nonexistent.png", 1, 1, 3, 3));
    }



    #endregion
}
