﻿using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDocumentsApi.Word.Interfaces;

namespace OfficeDocumentsApi.Word.DataClasses
{
    public class Body : IBody
    {
        internal DocumentFormat.OpenXml.Wordprocessing.Body Element { get; }
        public List<Paragraph> Paragraphs { get; } = new List<Paragraph>();

        public Body(DocumentFormat.OpenXml.Wordprocessing.Body element)
        {
            Element = element;
            foreach (var child in element.ChildElements)
            {
                switch (child)
                {
                    case DocumentFormat.OpenXml.Wordprocessing.Paragraph p:
                        Paragraphs.Add(new Paragraph(p));
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

        public string GetAllTexts()
        {
            return string.Join("\n", Paragraphs.Select(x => x.GetTexts()).Where(z => !string.IsNullOrEmpty(z)).ToArray());
        }
    }
}
