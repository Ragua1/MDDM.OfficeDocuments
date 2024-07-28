# OfficeDocuments.Excel
## Overview
**OfficeDocuments.Excel** is a sophisticated C# library designed to streamline the manipulation and creation of Excel documents using the OpenXml library. 
It simplifies the complexities associated with OpenXml by providing robust utilities for creating, modifying, and merging styles. 
These styles encompass text fonts, cell fills, and borders.  

The library allows for continuous creation, modification, and merging of styles, with the document being saved only upon invoking the Close or Dispose function. 
This ensures that all modifications are applied before the document is finalized.

## Usage
### Creating a new Excel document to stream
```csharp
using System.IO;
using OfficeDocuments.Excel;

var stream = new MemoryStream();
var spreadsheet = Spreadsheet.CreateDocument(stream);

// Add a worksheet
var worksheet = spreadsheet.AddWorksheet("Sheet1");

// Add some data to the worksheet
worksheet.Cells["A1"].Value = "Hello";
worksheet.Cells["B1"].Value = "World";

// Save and close the document
spreadsheet.Close();
```

### Opening an existing Excel document from a stream
```csharp
using System.IO;
using OfficeDocuments.Excel;

var stream = new FileStream("path/to/existing/file.xlsx", FileMode.Open, FileAccess.Read);
var spreadsheet = Spreadsheet.OpenDocument(stream, isEditable: true);

// Get the first worksheet
var worksheet = spreadsheet.GetWorksheet("Sheet1");

// Read some data from the worksheet
var valueA1 = worksheet.Cells["A1"].Value;
var valueB1 = worksheet.Cells["B1"].Value;

// Close the document
spreadsheet.Close();
```

### Creating a new Excel document to file
```csharp
using OfficeDocuments.Excel;

var filePath = "path/to/new/file.xlsx";
var spreadsheet = new Spreadsheet(filePath, createNew: true);

// Add a worksheet
var worksheet = spreadsheet.AddWorksheet("Sheet1");

// Add some data to the worksheet
worksheet.Cells["A1"].Value = "Hello";
worksheet.Cells["B1"].Value = "World";

// Save and close the document
spreadsheet.Close();
```

### Opening an existing Excel document from a file
```csharp
using OfficeDocuments.Excel;

var filePath = "path/to/existing/file.xlsx";
var spreadsheet = new Spreadsheet(filePath, createNew: false);

// Get the first worksheet
var worksheet = spreadsheet.GetWorksheet("Sheet1");

// Read some data from the worksheet
var valueA1 = worksheet.Cells["A1"].Value;
var valueB1 = worksheet.Cells["B1"].Value;

// Close the document
spreadsheet.Close();
```

## Features
### Styles in a Document
The library allows for extensive styling of Excel documents. You can style cell alignment, borders, fill colors, fonts, and number formats. 
Below are examples demonstrating how to apply these styles.

**Alignment**:
You can set the alignment of cell content both horizontally and vertically.
```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Styles;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var alignment = new Alignment
{
    Horizontal = HorizontalAlignmentValues.Center,
    Vertical = VerticalAlignmentValues.Center
};

worksheet.Cells["A1"].Style.Alignment = alignment;
spreadsheet.Close();
```
**Border**:
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

worksheet.Cells["A1"].Style.Border = border;
spreadsheet.Close();
```
**Fill**:
You can set the background and foreground colors of cells, as well as the fill pattern.
```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Styles;
using System.Drawing;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var fill = new Fill(Color.Yellow, Color.Red, PatternValues.Solid);

worksheet.Cells["A1"].Style.Fill = fill;
spreadsheet.Close();
```
**Font**:
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

worksheet.Cells["A1"].Style.Font = font;
spreadsheet.Close();
```
**Number Format**:
```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Styles;

var spreadsheet = new Spreadsheet("path/to/file.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var numberingFormat = new NumberingFormat("0.00");

worksheet.Cells["A1"].Style.NumberingFormat = numberingFormat;
spreadsheet.Close()
```


## Contributing
Contributions are welcome! Please feel free to submit a pull request or open an issue if you encounter a bug or have a feature request.

## License
This project is licensed under the MIT License. See the LICENSE file for details.