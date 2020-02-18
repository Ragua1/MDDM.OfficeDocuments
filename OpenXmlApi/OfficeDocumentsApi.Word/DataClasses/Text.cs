namespace OfficeDocumentsApi.Word.DataClasses
{
    public class Text
    {
        internal DocumentFormat.OpenXml.Wordprocessing.Text Element { get; }
        public Text(DocumentFormat.OpenXml.Wordprocessing.Text element)
        {
            Element = element;
        }
    }
}
