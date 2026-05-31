using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeDocuments.Word.Interfaces;

namespace OfficeDocuments.Word;

public class Wordprocessing : IWordprocessing
{
    private readonly WordprocessingDocument _document;
    private bool _isEditable = true;


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
            //var body = mainPart.Document.AppendChild(new Body());
            //var para = body.AppendChild(new Paragraph());
            //var run = para.AppendChild(new Run());
            //run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
            //run.AppendChild(new Break{Type = BreakValues.Page});

            //body = mainPart.Document.AppendChild(new Body()); 
            //para = body.AppendChild(new Paragraph());
            //run = para.AppendChild(new Run());
            //run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
        }
        else
        {
            var mainPart = document.MainDocumentPart;
            var doc = mainPart.Document;
            //doc.Body


            //var parts = document.Parts;

            ;
            //document.
        }

        _document = document;
    }

    public IBody GetBody()
    {
        var doc = GetDocument();

        Body bodyElement;
        if (doc.Body == null)
        {
            bodyElement = new Body();
            doc.AppendChild(bodyElement);
        }
        else
        {
            bodyElement = doc.Body;
        }

        //var bodyElement = doc.Body ?? new Body();

        //return doc.Body != null ? new DataClasses.Body(doc.Body) : new DataClasses.Body();
        return new DataClasses.Body(bodyElement);
    }

    private Document GetDocument()
    {
        var mainPart = _document.MainDocumentPart;
        if (mainPart is null)
        {
            throw new InvalidOperationException("The word document does not contain a main document part.");
        }

        var wordDocument = mainPart.Document;
        if (wordDocument is null)
        {
            throw new InvalidOperationException("The word document does not contain a root document element.");
        }

        return wordDocument;
    }

    #region IDisposable implementation

    /// <summary>
    /// Save and close document
    /// </summary>
    public void Close(bool saveDocument = true)
    {
        if (_isEditable && saveDocument)
        {
            _document.Save();
        }
        _document.Dispose();
    }

    /// <summary>
    /// Close document resources
    /// </summary>
    public void Dispose()
    {
        using (_document)
        {
            Close();
        }
    }

    #endregion
}