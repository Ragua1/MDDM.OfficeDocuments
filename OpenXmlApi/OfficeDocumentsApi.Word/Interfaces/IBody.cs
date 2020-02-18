using System.Collections.Generic;
using OfficeDocumentsApi.Word.DataClasses;

namespace OfficeDocumentsApi.Word.Interfaces
{
    public interface IBody
    {
        List<IParagraph> Paragraphs { get; }
        IParagraph AddParagraph();
        string GetAllTexts();
    }
}