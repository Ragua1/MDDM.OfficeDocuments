namespace OfficeDocuments.Excel.Interfaces
{
    public interface IOpenXmlWrapper<out T>
    {
        T Element { get; }
    }
}