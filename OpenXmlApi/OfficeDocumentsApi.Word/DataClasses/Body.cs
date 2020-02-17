using System;
using System.Collections.Generic;
using OfficeDocumentsApi.Word.Interfaces;

namespace OfficeDocumentsApi.Word.DataClasses
{
    public class Body : IBody
    {
        internal readonly DocumentFormat.OpenXml.Wordprocessing.Body Element;
        public List<Paragraph> Paragraphs { get; } = new List<Paragraph>();

        public Body(DocumentFormat.OpenXml.Wordprocessing.Body element)
        {
            this.Element = element;
            foreach (var child in element.ChildElements)
            {
                switch (child)
                {
                    case DocumentFormat.OpenXml.Wordprocessing.Paragraph p:
                        Paragraphs.Add(new Paragraph(p));
                        break;
                    case null:
                        throw new ArgumentNullException();
                        break;
                    default:
                        break;
                }
            }
        }

        public IParagraph AddParagraph()
        { 
            var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            this.Element.AppendChild(paragraph);

            return new Paragraph(paragraph);
        }
    }
}