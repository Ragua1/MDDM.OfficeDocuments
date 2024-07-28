using System.Collections.Generic;
using OfficeDocuments.Word.Enums;

namespace OfficeDocuments.Word.Interfaces
{
    public interface IParagraph
    {
        IParagraph AddText(string text);
        IParagraph AddBreak(BreakType type);
        IEnumerable<IText> GetTextElements();
        string GetTexts();
    }
}