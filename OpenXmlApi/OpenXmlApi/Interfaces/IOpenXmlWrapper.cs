namespace OpenXmlApi.Interfaces
{
    public interface IOpenXmlWrapper<out T>
    {
        T Element { get; }
    }
}