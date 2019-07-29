using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlApi.Interfaces;
using OpenXmlApi.Styles;
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
            throw new NotImplementedException();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Spreadsheet"/> class.
        /// </summary>
        /// <param name="filepath">The filepath of document</param>
        /// <exception cref="DirectoryNotFoundException">Exception for not exist path of file</exception>
        public static Spreadsheet Create(string filepath)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Spreadsheet"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <exception cref="DirectoryNotFoundException">Exception for not exist path of file</exception>
        public static Spreadsheet Create(Stream stream)
        {
            throw new NotImplementedException();
        }

        private static Spreadsheet Open(SpreadsheetDocument document, bool isEditable)
        {
            throw new NotImplementedException();
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
            throw new NotImplementedException();
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
            throw new NotImplementedException();
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
            throw new NotImplementedException();
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
            throw new NotImplementedException();
        }

        private SpreadsheetLib.Stylesheet InitStylesheet()
        {
            throw new NotImplementedException();
        }
    }
}