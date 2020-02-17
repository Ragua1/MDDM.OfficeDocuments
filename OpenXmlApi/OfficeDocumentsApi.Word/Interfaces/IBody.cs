using System.Collections.Generic;
using OfficeDocumentsApi.Word.DataClasses;

namespace OfficeDocumentsApi.Word.Interfaces
{
    public interface IBody
    {
        List<Paragraph> Paragraphs { get; }
        IParagraph AddParagraph();
    }
}