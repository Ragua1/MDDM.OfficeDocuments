using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Schema;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlApi.Emums;
using OpenXmlApi.Interfaces;
using OpenXmlApi.Styles;
using Color = System.Drawing.Color;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlApi
{
    /// <summary>
    /// Class of Spreadsheet
    /// </summary>
    public class Spreadsheet : IDisposable
    {
        /// <summary>
        /// Collection of worksheet in document
        /// </summary>
        public readonly List<IWorksheet> Worksheets = new List<IWorksheet>();
        private readonly SpreadsheetDocument document;
        internal WorkbookPart WorkbookPart => this.document.WorkbookPart;
        private SpreadsheetLib.Sheets Sheets => this.document.WorkbookPart.Workbook.Sheets;
        internal WorkbookStylesPart WorkbookStylesPart => this.WorkbookPart.WorkbookStylesPart;
        private IStyle defaultStyle;
        private bool IsEditable = true;

        internal SpreadsheetLib.Stylesheet Stylesheet => this.WorkbookStylesPart.Stylesheet ?? InitStylesheet();

        private Spreadsheet(SpreadsheetDocument document, bool createNew)
        {
            // Create a spreadsheet document
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            this.document = document;

            if (createNew)
            {
                this.document.AddWorkbookPart();
                this.WorkbookPart.Workbook = new SpreadsheetLib.Workbook();

                // Add Sheets to the Workbook.
                this.WorkbookPart.Workbook.AppendChild(new SpreadsheetLib.Sheets());

                // Add the WorkbookStylesPart.
                this.WorkbookPart.AddNewPart<WorkbookStylesPart>();

                //Init Stylesheet
                InitStylesheet();
            }
            else
            {
                if (this.WorkbookPart?.Workbook == null)
                {
                    throw new XmlSchemaValidationException("The document is not valid!");
                }

                if (this.WorkbookPart.WorkbookStylesPart == null)
                {
                    // Add the WorkbookStylesPart.
                    this.WorkbookPart.AddNewPart<WorkbookStylesPart>();

                    //Init Stylesheet
                    InitStylesheet();
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Spreadsheet"/> class.
        /// </summary>
        /// <param name="filepath">The filepath of document</param>
        /// <exception cref="DirectoryNotFoundException">Exception for not exist path of file</exception>
        public static Spreadsheet Create(string filepath)
        {
            var document = SpreadsheetDocument.Create(Path.GetFullPath(filepath), SpreadsheetDocumentType.Workbook);

            return new Spreadsheet(document, true);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Spreadsheet"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <exception cref="DirectoryNotFoundException">Exception for not exist path of file</exception>
        public static Spreadsheet Create(Stream stream)
        {
            var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

            return new Spreadsheet(document, true);
        }

        private static Spreadsheet Open(SpreadsheetDocument document, bool isEditable)
        {
            var writer = new Spreadsheet(document, false)
            {
                IsEditable = isEditable
            };

            foreach (var worksheetPart in writer.WorkbookPart.WorksheetParts)
            {
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SpreadsheetLib.SheetData>();

                var worksheet = new Worksheet(writer, worksheetPart, sheetData);
                writer.Worksheets.Add(worksheet);
            }

            return writer;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Spreadsheet"/> class for existing file
        /// </summary>
        /// <param name="filepath">The filepath of document</param>
        /// <returns>Instance of <see cref="Spreadsheet"/></returns>
        public static Spreadsheet Open(string filepath)
        {
            return Open(SpreadsheetDocument.Open(Path.GetFullPath(filepath), true), true);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Spreadsheet"/> class for existing file
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="isEditable"></param>
        /// <returns>Instance of <see cref="Spreadsheet"/></returns>
        public static Spreadsheet Open(Stream stream, bool isEditable = true)
        {
            return Open(SpreadsheetDocument.Open(stream, isEditable), isEditable);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Spreadsheet"/> class for existing file or create new file
        /// </summary>
        /// <param name="filepath">The filepath of document</param>
        /// <returns>Instance of <see cref="Spreadsheet"/></returns>
        public static Spreadsheet OpenOrCreate(string filepath)
        {
            return File.Exists(filepath) ? Open(filepath) : Create(filepath);
        }

        /// <summary>
        /// Create worksheet and apply 'style'
        /// </summary>
        /// <param name="sheetName">Worksheet name</param>
        /// <param name="sheetStyle">Custom style for worksheet</param>
        /// <returns>Created worksheet</returns>
        public IWorksheet AddWorksheet(string sheetName = null, IStyle sheetStyle = null)
        {
            var sheetData = new SpreadsheetLib.SheetData();

            // Add a blank WorksheetPart.
            var worksheetPart = this.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new SpreadsheetLib.Worksheet(sheetData);

            string relationshipId = this.WorkbookPart.GetIdOfPart(worksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = this.Sheets.Elements<SpreadsheetLib.Sheet>().Any()
                ? this.Sheets.Elements<SpreadsheetLib.Sheet>().Select(s => s.SheetId.Value).Max() + 1
                : 1;

            // Append the new worksheet and associate it with the workbook.
            var sheet = new SpreadsheetLib.Sheet { Id = relationshipId, SheetId = sheetId, Name = sheetName ?? $"Sheet {sheetId}" };
            this.Sheets.Append(sheet);

            var worksheet = new Worksheet(this, worksheetPart, sheetData, this.defaultStyle?.CreateMergedStyle(sheetStyle));
            this.Worksheets.Add(worksheet);

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
        public IStyle CreateStyle(Font font = null, Fill fill = null, Border border = null, NumberingFormat numberFormat = null, Alignment alignment = null)
        {
            return CreateStyle(this.Stylesheet, font, fill, border, numberFormat, alignment);
        }

        /// <summary>
        /// Get worksheet by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Worksheet or null</returns>
        public IWorksheet GetWorksheet(string name)
        {
            var sheet = this.Sheets.Elements<SpreadsheetLib.Sheet>().FirstOrDefault(s => s.Name == name);
            if (sheet == null)
            {
                return null;
            }

            return this.Worksheets.FirstOrDefault(
                w => this.WorkbookPart.GetIdOfPart(((Worksheet)w).WorksheetPart) == sheet.Id
            );
        }

        public string[] GetWorksheetsName()
        {
            return this.Sheets.Elements<SpreadsheetLib.Sheet>().Select(s => s.Name.Value).ToArray();
        }

        /// <summary>
        /// Save and close document
        /// </summary>
        public void Close()
        {
            if (this.IsEditable)
            {
                this.WorkbookPart.Workbook.Save();
            }
            this.document.Close();
        }
        /// <summary>
        /// Close document resources
        /// </summary>
        public void Dispose()
        {
            Close();
        }

        private IStyle CreateStyle(SpreadsheetLib.Stylesheet stylesheet, Font font = null, Fill fill = null, Border border = null, NumberingFormat numberFormat = null, Alignment alignment = null)
        {
            return new Style(stylesheet ?? this.Stylesheet, font, fill, border, numberFormat, alignment);
        }

        private SpreadsheetLib.Stylesheet InitStylesheet()
        {
            var stylesheet = this.WorkbookStylesPart.Stylesheet = new SpreadsheetLib.Stylesheet();

            stylesheet.CellFormats = new SpreadsheetLib.CellFormats();
            stylesheet.Fills = new SpreadsheetLib.Fills(
                new SpreadsheetLib.Fill { PatternFill = new SpreadsheetLib.PatternFill { PatternType = SpreadsheetLib.PatternValues.None } },
                new SpreadsheetLib.Fill { PatternFill = new SpreadsheetLib.PatternFill { PatternType = SpreadsheetLib.PatternValues.Gray125 } }
            );

            this.defaultStyle = CreateStyle(
                stylesheet,
                new Font { FontSize = 11, Color = Color.Black, FontName = FontNameValues.Calibri },
                null,
                new Border()
            );

            stylesheet.CellStyleFormats = new SpreadsheetLib.CellStyleFormats(this.defaultStyle.Element.CloneNode(true));

            return stylesheet;
        }
    }
}