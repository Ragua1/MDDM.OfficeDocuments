using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeDocumentsApi
{
    public class Wordprocessing : IDisposable
    {
        private readonly WordprocessingDocument document;
        private bool IsEditable = true;


        public Wordprocessing(Stream stream, bool createNew) :
            this(createNew
                ? WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document)
                : WordprocessingDocument.Open(stream, true),
                createNew)
        { }
        public Wordprocessing(string filePath, bool createNew) :
            this(createNew
                ? WordprocessingDocument.Create(Path.GetFullPath(filePath), WordprocessingDocumentType.Document)
                : WordprocessingDocument.Open(Path.GetFullPath(filePath), true),
                createNew)
        { }


        protected internal Wordprocessing(WordprocessingDocument document, bool createNew)
        {
            if (createNew)
            {

            }
            else
            {

            }
        }

        #region IDisposable implementation

        /// <summary>
        /// Save and close document
        /// </summary>
        public void Close()
        {
            if (IsEditable)
            {
                //WorkbookPart.Workbook.Save();
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

        #endregion
    }
}