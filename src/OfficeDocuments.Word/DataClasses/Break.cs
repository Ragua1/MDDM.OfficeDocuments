namespace OfficeDocuments.Word.DataClasses;

public class Break
{
    internal DocumentFormat.OpenXml.Wordprocessing.Break Element { get; }
    public Break(DocumentFormat.OpenXml.Wordprocessing.Break element)
    {
        Element = element;
    }
}