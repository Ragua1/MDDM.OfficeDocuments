using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Schema;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlApi.DataClasses;
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
    public class Spreadsheet : IDisposable, ISpreadsheet
    {
        /// <summary>
        /// Collection of worksheet in document
        /// </summary>
        public readonly List<IWorksheet> Worksheets = new List<IWorksheet>();
        private readonly SpreadsheetDocument document;
        public WorkbookPart WorkbookPart => document.WorkbookPart;
        public SpreadsheetLib.Sheets Sheets => document.WorkbookPart.Workbook.Sheets;
        public WorkbookStylesPart WorkbookStylesPart => WorkbookPart.WorkbookStylesPart;
        private IStyle defaultStyle;
        private bool IsEditable = true;

        public SpreadsheetLib.Stylesheet Stylesheet => WorkbookStylesPart.Stylesheet ?? InitStylesheet();

        protected internal Spreadsheet(SpreadsheetDocument document, bool createNew)
        {
            this.document = document;

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
                    Worksheets.Add(worksheet);
                }
            }
        }

        public Spreadsheet(Stream stream, bool createNew) :
            this(createNew
                ? SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)
                : SpreadsheetDocument.Open(stream, true),
                createNew)
        {
            // Create a spreadsheet document
            // By default, AutoSave = true, Editable = true, and Type = xlsx.

            //document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

           
        }

        public Spreadsheet(string filePath, bool createNew) : 
            this(createNew
                ? SpreadsheetDocument.Create(Path.GetFullPath(filePath), SpreadsheetDocumentType.Workbook)
                : SpreadsheetDocument.Open(Path.GetFullPath(filePath), true),
                createNew) { }

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
            var worksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new SpreadsheetLib.Worksheet(sheetData);

            var relationshipId = WorkbookPart.GetIdOfPart(worksheetPart);

            // Get a unique ID for the new worksheet.
            var sheetId = Sheets.Elements<SpreadsheetLib.Sheet>().Any()
                ? Sheets.Elements<SpreadsheetLib.Sheet>().Select(s => s.SheetId.Value).Max() + 1
                : 1;

            // Append the new worksheet and associate it with the workbook.
            var sheet = new SpreadsheetLib.Sheet { Id = relationshipId, SheetId = sheetId, Name = sheetName ?? $"Sheet {sheetId}" };
            Sheets.Append(sheet);

            var worksheet = new Worksheet(this, worksheetPart, sheetData, defaultStyle?.CreateMergedStyle(sheetStyle));
            Worksheets.Add(worksheet);

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
            return CreateStyle(Stylesheet, font, fill, border, numberFormat, alignment);
        }

        /// <summary>
        /// Get worksheet by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Worksheet or null</returns>
        public IWorksheet GetWorksheet(string name)
        {
            var sheet = Sheets.Elements<SpreadsheetLib.Sheet>().FirstOrDefault(s => s.Name == name);
            if (sheet == null)
            {
                return null;
            }

            return Worksheets.FirstOrDefault(
                w => WorkbookPart.GetIdOfPart(((Worksheet)w).WorksheetPart) == sheet.Id
            );
        }

        public string[] GetWorksheetsName()
        {
            return Sheets.Elements<SpreadsheetLib.Sheet>().Select(s => s.Name.Value).ToArray();
        }

        /// <summary>
        /// Save and close document
        /// </summary>
        public void Close()
        {
            if (IsEditable)
            {
                WorkbookPart.Workbook.Save();
            }
            document.Close();
        }
        /// <summary>
        /// Close document resources
        /// </summary>
        public void Dispose()
        {
            using (document)
            {
                Close();
            }
        }

        public IStyle CreateStyle(SpreadsheetLib.Stylesheet stylesheet, Font font = null, Fill fill = null, Border border = null, NumberingFormat numberFormat = null, Alignment alignment = null)
        {
            return new Style(stylesheet ?? Stylesheet, font, fill, border, numberFormat, alignment);
        }

        public SpreadsheetLib.Stylesheet InitStylesheet()
        {
            var stylesheet = WorkbookStylesPart.Stylesheet = new SpreadsheetLib.Stylesheet();

            stylesheet.CellFormats = new SpreadsheetLib.CellFormats();
            stylesheet.Fills = new SpreadsheetLib.Fills(
                new SpreadsheetLib.Fill { PatternFill = new SpreadsheetLib.PatternFill { PatternType = SpreadsheetLib.PatternValues.None } },
                new SpreadsheetLib.Fill { PatternFill = new SpreadsheetLib.PatternFill { PatternType = SpreadsheetLib.PatternValues.Gray125 } }
            );

            defaultStyle = CreateStyle(
                stylesheet,
                new Font { FontSize = 11, Color = Color.Black, FontName = FontNameValues.Calibri },
                null,
                new Border()
            );

            stylesheet.CellStyleFormats = new SpreadsheetLib.CellStyleFormats(defaultStyle.Element.CloneNode(true));

            return stylesheet;
        }
    }
}