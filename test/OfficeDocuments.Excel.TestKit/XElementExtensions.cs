using System.Xml.Linq;

namespace OfficeDocuments.Excel.TestKit;

/// <summary>
/// Namespace- and order-normalized XML equality, for asserting on generated OOXML fragments
/// without depending on attribute or sibling ordering.
/// </summary>
public static class XElementExtensions
{
    extension(string xml)
    {
        public bool CompareXml(string otherXml)
        {
            ArgumentNullException.ThrowIfNull(xml);
            ArgumentNullException.ThrowIfNull(otherXml);

            return CompareXml(
                XDocument.Parse(xml).Root ?? throw new InvalidOperationException(),
                XDocument.Parse(otherXml).Root ?? throw new InvalidOperationException());
        }
    }

    private static bool CompareXml(XElement left, XElement right)
    {
        return XNode.DeepEquals(
            Normalize(left),
            Normalize(right));
    }

    private static XElement Normalize(XElement element)
    {
        // Namespace declarations are bookkeeping, not content: the same element written with a
        // default namespace and with an "x:" prefix must compare equal.
        var attributes = element.Attributes()
            .Where(attribute => !attribute.IsNamespaceDeclaration)
            .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal)
            .Select(attribute => new XAttribute(attribute));

        if (!element.HasElements)
        {
            if (element.IsEmpty || string.IsNullOrEmpty(element.Value))
            {
                return new XElement(element.Name, attributes);
            }

            return new XElement(element.Name, attributes, element.Value);
        }

        var children = element.Elements()
            .Select(Normalize)
            .OrderBy(child => child.Name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(child => child.Name.LocalName, StringComparer.Ordinal)
            .ThenBy(child => child.ToString(SaveOptions.DisableFormatting), StringComparer.Ordinal);

        return new XElement(element.Name, attributes, children);
    }
}