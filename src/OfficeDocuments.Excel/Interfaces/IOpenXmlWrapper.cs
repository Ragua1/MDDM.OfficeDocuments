using System.ComponentModel;

namespace OfficeDocuments.Excel.Interfaces;

public interface IOpenXmlWrapper<out T>
{
    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This property exposes the raw OpenXml element. Prefer using the typed interface API instead.")]
    T Element { get; }
}