using OfficeDocumentsApi.Word.Enums;

namespace OfficeDocumentsApi.Word.Interfaces
{
    public interface IParagraph
    {
        IParagraph AddText(string text);
        IParagraph AddBreak(BreakType type);
        string GetTexts();
    }
}