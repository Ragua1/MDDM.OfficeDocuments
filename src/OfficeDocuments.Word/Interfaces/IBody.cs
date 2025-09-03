using System.Collections.Generic;

namespace OfficeDocuments.Word.Interfaces;

public interface IBody
{
    List<IParagraph> Paragraphs { get; }
    IParagraph AddParagraph();
    string GetAllTexts();
}