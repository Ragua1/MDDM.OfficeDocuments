using System.Xml.Linq;

namespace OfficeDocuments.Excel.Extensions;

public static class XElementExtensions
{
    /// <summary>
    /// Compare two xml by content
    /// </summary>
    public static bool CompareXml(this string xml1, string xml2)
    {
        ArgumentNullException.ThrowIfNull(xml1, nameof(xml1));
        ArgumentNullException.ThrowIfNull(xml2, nameof(xml2));
        
        return CompareXml(
            XDocument.Parse(xml1).Root ?? throw new InvalidOperationException(), 
            XDocument.Parse(xml2).Root ?? throw new InvalidOperationException()
            );
    }

    private static bool CompareXml(this XElement elm1, XElement elm2)
    {
        return XNode.DeepEquals(
            Normalize(elm1),
            Normalize(elm2)
        );
    }

    private static XElement Normalize(this XElement element)
    {
        if (element.HasElements)
        {
            return new XElement(element.Name, element.Attributes().Where(a => a.Name.Namespace == XNamespace.Xmlns)
                .OrderBy(a => a.Name.ToString()), element.Elements().OrderBy(a => a.Name.ToString())
                .Select(Normalize));
        }

        if (element.IsEmpty || string.IsNullOrEmpty(element.Value))
        {
            return new XElement(element.Name, element.Attributes()
                .OrderBy(a => a.Name.ToString()));
        }

        return new XElement(element.Name, element.Attributes()
            .OrderBy(a => a.Name.ToString()), element.Value);
    }
}