# OfficeDocuments

This library provides convenient APIs to work with Office document formats (Excel and Word) using the OpenXml SDK.

## OfficeDocuments.Excel
### Overview
**OfficeDocuments.Excel** is a sophisticated C# library designed to streamline the manipulation and creation of Excel documents using the OpenXml library. 
It simplifies the complexities associated with OpenXml by providing robust utilities for creating, modifying, and merging styles. 
These styles encompass text fonts, cell fills, and borders.  

The library allows for continuous creation, modification, and merging of styles, with the document being saved only upon invoking the Close or Dispose function. 
This ensures that all modifications are applied before the document is finalized.

### Usage
#### Creating a new Excel document to stream
```csharp
using System.IO;
using OfficeDocuments.Excel;

var stream = new MemoryStream();
var spreadsheet = Spreadsheet.CreateDocument(stream);

// Add a worksheet
var worksheet = spreadsheet.AddWorksheet("Sheet1");

// Add some data to the worksheet
var row = worksheet.AddRow();
row.AddCell("Hello");
row.AddCell("World");

// Save and close the document
spreadsheet.Close();
```

#### Opening an existing Excel document from a stream
```csharp
using System.IO;
using OfficeDocuments.Excel;

var stream = new FileStream("path/to/existing/file.xlsx", FileMode.Open, FileAccess.Read);
var spreadsheet = Spreadsheet.OpenDocument(stream, isEditable: true);

// Get the first worksheet
var worksheet = spreadsheet.GetWorksheet("Sheet1");

// Read some data from the worksheet
var valueA1 = worksheet.GetCellByReference("A1").Value;
var valueB1 = worksheet.GetCellByReference("B1").Value;

// Close the document
spreadsheet.Close();
```

#### Creating a new Excel document to file
```csharp
using OfficeDocuments.Excel;

var filePath = "path/to/new/file.xlsx";
var spreadsheet = new Spreadsheet(filePath, createNew: true);

// Add a worksheet
var worksheet = spreadsheet.AddWorksheet("Sheet1");

// Add some data to the worksheet
var row = worksheet.AddRow();
row.AddCell("Hello");
row.AddCell("World");

// Save and close the document
spreadsheet.Close();
```

#### Opening an existing Excel document from a file
```csharp
using OfficeDocuments.Excel;

var filePath = "path/to/existing/file.xlsx";
var spreadsheet = new Spreadsheet(filePath, createNew: false);

// Get the first worksheet
var worksheet = spreadsheet.GetWorksheet("Sheet1");

// Read some data from the worksheet
var valueA1 = worksheet.GetCellByReference("A1").Value;
var valueB1 = worksheet.GetCellByReference("B1").Value;

// Close the document
spreadsheet.Close();
```

### Features
#### Styles in a Document
The library allows for extensive styling of Excel documents. You can style cell alignment, borders, fill colors, fonts, and number formats. 
Below are examples demonstrating how to apply these styles.

##### Alignment
You can set the alignment of cell content both horizontally and vertically.
```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Styles;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var alignment = new Alignment
{
    Horizontal = HorizontalAlignmentValues.General,
    Vertical = VerticalAlignmentValues.Justify
};

var style1 = spreadsheet.CreateStyle(alignment: alignment);

worksheet.AddRow(style1).AddCell("Hello");
spreadsheet.Close();
```

##### Border
You can set the border of a cell to a specific color, style, and thickness.
```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Styles;
using System.Drawing;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var border = new Border
{
    LeftBorder = new BorderStyle { Style = BorderStyleValues.Thin, Color = Color.Black },
    RightBorder = new BorderStyle { Style = BorderStyleValues.Thin, Color = Color.Black },
    TopBorder = new BorderStyle { Style = BorderStyleValues.Thin, Color = Color.Black },
    BottomBorder = new BorderStyle { Style = BorderStyleValues.Thin, Color = Color.Black }
};

var style1 = spreadsheet.CreateStyle(border: border);

worksheet.AddRow(style1).AddCell("Hello");
spreadsheet.Close();
```

##### Fill
You can set the background and foreground colors of cells, as well as the fill pattern.
```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Styles;
using System.Drawing;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var fill = new Fill(Color.Yellow, Color.Red, PatternValues.Solid);

var style1 = spreadsheet.CreateStyle(fill: fill);

worksheet.AddRow(style1).AddCell("Hello");
spreadsheet.Close();
```

##### Font
You can customize the font of cell text, including its size, color, and style (bold, italic, underline).
```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Styles;
using System.Drawing;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var font = new Font
{
    Bold = true,
    Italic = true,
    FontSize = 12,
    FontName = FontNameValues.Calibri,
    Color = Color.Blue
};

var style1 = spreadsheet.CreateStyle(font: font);

worksheet.AddRow(style1).AddCell("Hello");
spreadsheet.Close();
```

##### Number Format
You can format numeric values displayed in cells using number formats.
```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Styles;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var numberingFormat = new NumberingFormat("0.00");

var style1 = spreadsheet.CreateStyle(numberingFormat: numberingFormat);

worksheet.AddRow(style1).AddCell(5.51);
spreadsheet.Close();
```

#### Advanced Features

##### Merging Styles
You can combine multiple styles using the CreateMergedStyle method.
```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Styles;
using System.Drawing;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

// Create individual styles
var boldFont = spreadsheet.CreateStyle(new Font { Bold = true });
var blueFont = spreadsheet.CreateStyle(new Font { Color = Color.Blue });
var border = spreadsheet.CreateStyle(border: new Border(BorderStyleValues.Medium));

// Merge styles together
var combinedStyle = boldFont.CreateMergedStyle(blueFont).CreateMergedStyle(border);

// Apply the combined style
worksheet.AddRow(combinedStyle).AddCell("Bold Blue Text with Border");
spreadsheet.Close();
```

##### Working with Cell Ranges
You can add cells to specific ranges in a worksheet.

```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Styles;
using System.Drawing;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var style = spreadsheet.CreateStyle(
    new Font { FontSize = 20, Color = Color.Blue, FontName = FontNameValues.Tahoma }
);

// Add a cell spanning multiple columns and rows (from column 3 to column 6, in row 2)
var cell = worksheet.AddCellOnRange(3, 6, 2, style);
cell.SetValue("Header Text Spanning Multiple Columns");

spreadsheet.Close();
```

##### Adding Formulas
You can add cells with Excel formulas.

```csharp
using OfficeDocuments.Excel;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

// Add data rows
var row1 = worksheet.AddRow();
row1.AddCell(100);
row1.AddCell(200);
row1.AddCell(300);

var row2 = worksheet.AddRow();
row2.AddCell(400);
row2.AddCell(500);
row2.AddCell(600);

// Add a sum formula in the next row
var sumRow = worksheet.AddRow();
sumRow.AddCell("Total:");
sumRow.AddCellWithFormula("SUM(A1:C2)"); // Will calculate the sum of all values

spreadsheet.Close();
```

##### Setting Column Width
You can customize the width of columns in the worksheet.

```csharp
using OfficeDocuments.Excel;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

// Set width of column A (index 1)
worksheet.SetColumnWidth(1, 20);

// Set width of column B (index 2)
worksheet.SetColumnWidth(2, 15);

spreadsheet.Close();
```

##### Working with In-Memory Spreadsheets
You can create and manipulate Excel documents entirely in memory without writing to disk.

```csharp
using System.IO;
using OfficeDocuments.Excel;

// Create a new document in memory
var memory = new MemoryStream();
using (var spreadsheet = Spreadsheet.CreateDocument(memory))
{
    var worksheet = spreadsheet.AddWorksheet("Sheet1");
    
    // Add data
    var cell = worksheet.AddCell("Sample Text");
    
    // document is saved to memory stream when Close() is called
}

// Rewind the stream to read it
memory.Position = 0;

// Open the in-memory document
using (var spreadsheet = Spreadsheet.OpenDocument(memory))
{
    var worksheet = spreadsheet.GetWorksheet("Sheet1");
    var cell = worksheet.GetCellByReference("A1");
    // cell.Value will be "Sample Text"
}
```

## OfficeDocuments.Word
The Word portion of the library provides similar functionality for Word documents. Documentation for this module will be expanded in the future.

## Contributing
Contributions are welcome! Please feel free to submit a pull request or open an issue if you encounter a bug or have a feature request.

## License
Mode info in file LICENSE.md