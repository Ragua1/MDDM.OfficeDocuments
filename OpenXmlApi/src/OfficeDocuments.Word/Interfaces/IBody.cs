using System.Collections.Generic;
using OfficeDocuments.Word.DataClasses;

namespace OfficeDocuments.Word.Interfaces
{
    public interface IBody
    {
        List<IParagraph> Paragraphs { get; }
        IParagraph AddParagraph();
        string GetAllTexts();
    }
}