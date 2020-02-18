﻿using OfficeDocumentsApi.Word.Interfaces;

namespace OfficeDocumentsApi.Word.DataClasses
{
    public class Text : IText
    {
        public string TextValue
        {
            get => Element.Text;
            set => Element.Text = value;
        }

        internal DocumentFormat.OpenXml.Wordprocessing.Text Element { get; }
        public Text(DocumentFormat.OpenXml.Wordprocessing.Text element)
        {
            Element = element;
        }
    }
}
