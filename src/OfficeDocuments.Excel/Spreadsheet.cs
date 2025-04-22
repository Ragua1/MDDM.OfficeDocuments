using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Schema;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.DataClasses;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;

using Color = System.Drawing.Color;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;
using Alignment = OfficeDocuments.Excel.Styles.Alignment;
using Border = OfficeDocuments.Excel.Styles.Border;
using Fill = OfficeDocuments.Excel.Styles.Fill;
using Font = OfficeDocuments.Excel.Styles.Font;
using NumberingFormat = OfficeDocuments.Excel.Styles.NumberingFormat;
using Worksheet = OfficeDocuments.Excel.DataClasses.Worksheet;

namespace OfficeDocuments.Excel
{

    /// <summary>
    /// Class of Spreadsheet
    /// </summary>
    public class Spreadsheet : ISpreadsheet, IDisposable
    {
        /// <summary>
        /// Collection of worksheet in document
        /// </summary>
        private readonly List<IWorksheet> _worksheets = new();

        /// <summary>
        /// Gets the collection of worksheets in the document
        /// </summary>
        public IReadOnlyList<IWorksheet> Worksheets => _worksheets.AsReadOnly();

        private readonly SpreadsheetDocument _document;
        private IStyle? _defaultStyle = null;
        private readonly bool _isEditable;
        private bool _disposed = false;

        public WorkbookPart WorkbookPart => _document.WorkbookPart ?? throw new InvalidOperationException();
        public SpreadsheetLib.Sheets Sheets => _document.WorkbookPart?.Workbook.Sheets ?? throw new InvalidOperationException();
        public WorkbookStylesPart WorkbookStylesPart => WorkbookPart.WorkbookStylesPart ?? throw new InvalidOperationException();
        public SpreadsheetLib.Stylesheet Stylesheet => WorkbookStylesPart.Stylesheet ?? InitStylesheet();

        private Spreadsheet(SpreadsheetDocument document, bool createNew, bool isEditable = true)
        {
            this._document = document;
            this._isEditable = isEditable;

            if (createNew)
            {
                document.AddWorkbookPart();
                WorkbookPart.Workbook = new SpreadsheetLib.Workbook();

                // Add Sheets to the Workbook.
                WorkbookPart.Workbook.AppendChild(new SpreadsheetLib.Sheets());

                // Add the WorkbookStylesPart.
                WorkbookPart.AddNewPart<WorkbookStylesPart>();

                //Init Stylesheet
                InitStylesheet();
            }
            else
            {
                if (WorkbookPart?.Workbook == null)
                {
                    throw new XmlSchemaValidationException("The document is not valid!");
                }

                if (WorkbookPart.WorkbookStylesPart == null)
                {
                    // Add the WorkbookStylesPart.
                    WorkbookPart.AddNewPart<WorkbookStylesPart>();

                    //Init Stylesheet
                    InitStylesheet();
                }

                foreach (var worksheetPart in WorkbookPart.WorksheetParts)
                {
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SpreadsheetLib.SheetData>();

                    var worksheet = new Worksheet(this, worksheetPart, sheetData);
                    _worksheets.Add(worksheet);
                }
            }
        }

        public Spreadsheet(Stream stream, bool createNew = false) :
            this(createNew
                ? SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)
                : SpreadsheetDocument.Open(stream, true),
                createNew)
        {
            // Create a spreadsheet document
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
        }
        public Spreadsheet(string filePath, bool createNew = false) :
            this(createNew
                ? SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook)
                : SpreadsheetDocument.Open(filePath, true),
                createNew) { }
        
        public static ISpreadsheet CreateDocument(Stream stream)
        {
            return new Spreadsheet(
                SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook), 
                true
            );
        }
        public static ISpreadsheet OpenDocument(Stream stream, bool isEditable = true)
        {
            return new Spreadsheet(
                SpreadsheetDocument.Open(stream, isEditable), 
                false, 
                isEditable
            );
        }

        /// <summary>
        /// Create worksheet and apply 'style'
        /// </summary>
        /// <param name="sheetName">Worksheet name</param>
        /// <param name="sheetStyle">Custom style for worksheet</param>
        /// <returns>Created worksheet</returns>
        public IWorksheet AddWorksheet(string? sheetName = null, IStyle? sheetStyle = null)
        {
            var sheetData = new SpreadsheetLib.SheetData();

            // Add a blank WorksheetPart.
            var worksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new SpreadsheetLib.Worksheet(sheetData);

            var relationshipId = WorkbookPart.GetIdOfPart(worksheetPart);

            // Get a unique ID for the new worksheet.
            var sheetId = Sheets.Elements<SpreadsheetLib.Sheet>().Any()
                ? Sheets.Elements<SpreadsheetLib.Sheet>().Select(s => s.SheetId.Value).Max() + 1
                : 1;

            // Generate a default sheet name if none provided
            var finalSheetName = sheetName ?? $"Sheet {sheetId}";

            // Append the new worksheet and associate it with the workbook.
            var sheet = new SpreadsheetLib.Sheet { Id = relationshipId, SheetId = sheetId, Name = finalSheetName };
            Sheets.Append(sheet);

            var worksheet = new Worksheet(this, worksheetPart, sheetData, _defaultStyle?.CreateMergedStyle(sheetStyle));
            _worksheets.Add(worksheet);

            return worksheet;
        }

        /// <summary>
        /// Create custom style
        /// </summary>
        /// <param name="font">Custom font styling</param>
        /// <param name="fill">Custom fill styling</param>
        /// <param name="border">Custom border styling</param>
        /// <param name="numberFormat">Custom number format styling</param>
        /// <param name="alignment">Custom alignment styling</param>
        /// <returns>Created style</returns>
        public IStyle CreateStyle(Font? font = null, Fill? fill = null, Border? border = null, NumberingFormat? numberFormat = null, Alignment? alignment = null)
        {
            return CreateStyle(Stylesheet, font, fill, border, numberFormat, alignment);
        }

        /// <summary>
        /// Creates a style with the specified properties
        /// </summary>
        /// <param name="stylesheet">The stylesheet to use</param>
        /// <param name="font">Custom font styling</param>
        /// <param name="fill">Custom fill styling</param>
        /// <param name="border">Custom border styling</param>
        /// <param name="numberFormat">Custom number format styling</param>
        /// <param name="alignment">Custom alignment styling</param>
        /// <returns>The created style</returns>
        /// <exception cref="ArgumentNullException">Thrown when stylesheet is null</exception>
        public IStyle CreateStyle(SpreadsheetLib.Stylesheet stylesheet, Font? font = null, Fill? fill = null, Border? border = null, NumberingFormat? numberFormat = null, Alignment? alignment = null)
        {
            if (stylesheet == null)
            {
                throw new ArgumentNullException(nameof(stylesheet), "Stylesheet cannot be null");
            }
            
            return new Style(stylesheet, font, fill, border, numberFormat, alignment);
        }

        /// <summary>
        /// Get worksheet by name
        /// </summary>
        /// <param name="name">The name of the worksheet to retrieve</param>
        /// <returns>Worksheet if found, null otherwise</returns>
        public IWorksheet? GetWorksheet(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return null;
            }

            var sheet = Sheets.Elements<SpreadsheetLib.Sheet>().FirstOrDefault(s => s.Name == name);
            if (sheet == null)
            {
                return null;
            }

            return _worksheets.FirstOrDefault(
                w => WorkbookPart.GetIdOfPart(((Worksheet)w).WorksheetPart) == sheet.Id
            );
        }

        /// <summary>
        /// Adds a table to the specified worksheet
        /// </summary>
        /// <param name="worksheetName">The name of the worksheet</param>
        /// <param name="startCell">The starting cell of the table</param>
        /// <param name="endCell">The ending cell of the table</param>
        /// <param name="columnsName">The names of the table columns</param>
        /// <exception cref="ArgumentException">Thrown when worksheet cannot be found or table definition is invalid</exception>
        /// <exception cref="ArgumentNullException">Thrown when required parameters are null</exception>
        public void AddTable(string worksheetName, ICell startCell, ICell endCell, List<string> columnsName)
        {
            // Validate parameters
            if (string.IsNullOrEmpty(worksheetName))
            {
                throw new ArgumentNullException(nameof(worksheetName), "Worksheet name cannot be null or empty");
            }

            if (startCell == null)
            {
                throw new ArgumentNullException(nameof(startCell), "Start cell cannot be null");
            }

            if (endCell == null)
            {
                throw new ArgumentNullException(nameof(endCell), "End cell cannot be null");
            }

            if (columnsName == null)
            {
                throw new ArgumentNullException(nameof(columnsName), "Column names list cannot be null");
            }

            if (!columnsName.Any())
            {
                throw new ArgumentException("Column names list cannot be empty", nameof(columnsName));
            }

            if (columnsName.Any(string.IsNullOrWhiteSpace))
            {
                throw new ArgumentException("Table column names cannot be null or empty", nameof(columnsName));
            }

            if (startCell.RowIndex > endCell.RowIndex || startCell.ColumnIndex > endCell.ColumnIndex)
            {
                throw new ArgumentException("Invalid table definition: start cell must be before end cell");
            }

            // Find worksheet
            var sheetId = Sheets.Elements<SpreadsheetLib.Sheet>().FirstOrDefault(s => s.Name == worksheetName)?.Id;
            if (string.IsNullOrEmpty(sheetId))
            {
                throw new ArgumentException($"Cannot find worksheet with name '{worksheetName}'", nameof(worksheetName));
            }

            var wsp = WorkbookPart.WorksheetParts.FirstOrDefault(w => WorkbookPart.GetIdOfPart(w) == sheetId);
            if (wsp == null)
            {
                throw new ArgumentException($"Cannot find worksheet part for '{worksheetName}'", nameof(worksheetName));
            }

            // Create table
            var tablesCount = wsp.TableDefinitionParts.Count();
            var tableIndex = tablesCount + 1;
            var tableName = $"Table{tableIndex}";
            var tableDisplayName = $"Table{tableIndex}";
            var tableRid = $"rId{tableIndex}";

            var table = new SpreadsheetLib.Table
            {
                Reference = $"{startCell.CellReference}:{endCell.CellReference}",
                TableColumns = new SpreadsheetLib.TableColumns(),
                Id = (uint)tableIndex,
                Name = tableName,
                DisplayName = tableDisplayName
            };

            // Add columns
            uint columnId = 1;
            foreach (var columnName in columnsName)
            {
                var tableColumn = new SpreadsheetLib.TableColumn
                {
                    Name = columnName,
                    Id = columnId++
                };

                table.TableColumns.AppendChild(tableColumn);
            }

            // Add table to worksheet
            var tdp = wsp.AddNewPart<TableDefinitionPart>(tableRid);
            tdp.Table = table;

            // Add table parts if needed
            var existingTableParts = wsp.Worksheet.GetFirstChild<SpreadsheetLib.TableParts>();
            if (existingTableParts == null)
            {
                existingTableParts = new SpreadsheetLib.TableParts();
                wsp.Worksheet.Append(existingTableParts);
            }

            var tablePart = new SpreadsheetLib.TablePart { Id = tableRid };
            existingTableParts.Append(tablePart);
        }

        public IEnumerable<string> GetWorksheetsName()
        {
            return Sheets.Elements<SpreadsheetLib.Sheet>().Select(s => s.Name.Value).ToArray();
        }

        /// <summary>
        /// Save and close document
        /// </summary>
        public void Close()
        {
            if (_isEditable)
            {
                WorkbookPart.Workbook.Save();
            }
            
            if (_document != null && !_disposed)
            {
                _document.Dispose();
            }
            
            _disposed = true;
        }
        
        /// <summary>
        /// Close document resources
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        
        /// <summary>
        /// Dispose pattern implementation
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;
                
            if (disposing)
            {
                Close();
            }
            
            _disposed = true;
        }
        
        /// <summary>
        /// Finalizer
        /// </summary>
        ~Spreadsheet()
        {
            Dispose(false);
        }

        public SpreadsheetLib.Stylesheet InitStylesheet()
        {
            var stylesheet = WorkbookStylesPart.Stylesheet = new SpreadsheetLib.Stylesheet();

            stylesheet.CellFormats = new SpreadsheetLib.CellFormats();
            stylesheet.Fills = new SpreadsheetLib.Fills(
                new SpreadsheetLib.Fill { PatternFill = new SpreadsheetLib.PatternFill { PatternType = SpreadsheetLib.PatternValues.None } },
                new SpreadsheetLib.Fill { PatternFill = new SpreadsheetLib.PatternFill { PatternType = SpreadsheetLib.PatternValues.Gray125 } }
            );

            _defaultStyle = CreateStyle(
                stylesheet,
                new Font { FontSize = 11, Color = Color.Black, FontName = FontNameValues.Calibri },
                null,
                new Border()
            );

            stylesheet.CellStyleFormats = new SpreadsheetLib.CellStyleFormats(_defaultStyle.Element.CloneNode(true));

            return stylesheet;
        }
    }
}