using System;
using System.Collections.Generic;
using System.Text;
using OfficeDocumentsApi.Word.Enums;
using OfficeDocumentsApi.Word.Interfaces;

namespace OfficeDocumentsApi.Word.DataClasses
{
    public class Paragraph : IParagraph
    {
        internal DocumentFormat.OpenXml.Wordprocessing.Paragraph Element { get; }

        public List<Run> RunList { get; }

        public Paragraph(DocumentFormat.OpenXml.Wordprocessing.Paragraph element)
        {
            RunList = new List<Run>();
            Element = element;
            foreach (var child in element.ChildElements)
            {
                switch (child)
                {
                    case DocumentFormat.OpenXml.Wordprocessing.Run run:
                        RunList.Add(new Run(run));
                        break;
                }
            }
        }

        public IParagraph AddText(string text)
        {
            var runElement = new DocumentFormat.OpenXml.Wordprocessing.Run();
            runElement.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(text));

            Element.AppendChild(runElement);
            RunList.Add(new Run(runElement));

            return this;
        }

        public IParagraph AddBreak(BreakType type)
        {
            DocumentFormat.OpenXml.Wordprocessing.BreakValues breakValue;
            switch (type)
            {
                case BreakType.Page:
                    breakValue = DocumentFormat.OpenXml.Wordprocessing.BreakValues.Page;
                    break;
                case BreakType.Column:
                    breakValue = DocumentFormat.OpenXml.Wordprocessing.BreakValues.Column;
                    break;
                case BreakType.TextWrapping:
                    breakValue = DocumentFormat.OpenXml.Wordprocessing.BreakValues.TextWrapping;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(type), type, null);
            }

            var runElement = new DocumentFormat.OpenXml.Wordprocessing.Run();
            runElement.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Break
            {
                Type = breakValue,
            });
            Element.AppendChild(runElement);
            RunList.Add(new Run(runElement));

            return this;
        }

        public string GetTexts()
        {
            var builder = new StringBuilder();

            foreach (var run in RunList)
            {
                foreach (var child in run.Element.ChildElements)
                {
                    switch (child)
                    {
                        case DocumentFormat.OpenXml.Wordprocessing.Text textElement:
                            var text = textElement.Text.Trim();
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                builder.Append(text);
                            }
                            break;
                    }
                }
            }

            return builder.ToString();
        }
    }
}