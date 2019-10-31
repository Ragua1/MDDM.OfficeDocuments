using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
                // Add a main document part. 
                var mainPart = document.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());
                var para = body.AppendChild(new Paragraph());
                var run = para.AppendChild(new Run());
                run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
            }
            else
            {

            }

            this.document = document;
        }

        #region IDisposable implementation

        /// <summary>
        /// Save and close document
        /// </summary>
        public void Close()
        {
            if (IsEditable)
            {
                document.Save();
            }
            document?.Close();
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