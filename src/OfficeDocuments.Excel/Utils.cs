using System.Collections.Generic;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using OfficeDocuments.Excel.Styles;

namespace OfficeDocuments.Excel;

/// <summary>
/// Class of utilities
/// </summary>
public static class Utils
{
    /// <summary>
    /// Convert color from 'System.Drawing.Color' to argb hex representation
    /// </summary>
    /// <param name="c"></param>
    /// <returns></returns>
    public static string ArgbHexConverter(System.Drawing.Color c)
    {
        return $"{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}";
    }

    /// <summary>
    /// Create new font by merging two fonts
    /// </summary>
    public static Font MergeFonts(DocumentFormat.OpenXml.Spreadsheet.Font font1, DocumentFormat.OpenXml.Spreadsheet.Font font2)
    {
        if (font1 == null) throw new ArgumentNullException(nameof(font1));
        if (font2 == null) throw new ArgumentNullException(nameof(font2));

        var mergedXml = MergeXmlElements(font1, font2);
        return new Font(new DocumentFormat.OpenXml.Spreadsheet.Font(mergedXml));
    }

    /// <summary>
    /// Create new fill by merging two fills
    /// </summary>
    public static Fill MergeFills(DocumentFormat.OpenXml.Spreadsheet.Fill fill1, DocumentFormat.OpenXml.Spreadsheet.Fill fill2)
    {
        if (fill1 == null) throw new ArgumentNullException(nameof(fill1));
        if (fill2 == null) throw new ArgumentNullException(nameof(fill2));

        var mergedXml = MergeXmlElements(fill1, fill2);
        return new Fill(new DocumentFormat.OpenXml.Spreadsheet.Fill(mergedXml));
    }

    /// <summary>
    /// Create new border by merging two borders
    /// </summary>
    public static Border MergeBorders(DocumentFormat.OpenXml.Spreadsheet.Border border1, DocumentFormat.OpenXml.Spreadsheet.Border border2)
    {
        if (border1 == null) throw new ArgumentNullException(nameof(border1));
        if (border2 == null) throw new ArgumentNullException(nameof(border2));

        var mergedXml = MergeXmlElements(border1, border2);
        return new Border(new DocumentFormat.OpenXml.Spreadsheet.Border(mergedXml));
    }

    /// <summary>
    /// Merges two OpenXml elements, with the second element's properties taking precedence
    /// </summary>
    private static string MergeXmlElements<T>(T element1, T element2) where T : OpenXmlElement
    {
        try
        {
            var xml1 = XDocument.Parse(element1.OuterXml);
            var xml2 = XDocument.Parse(element2.OuterXml);
            return MergeXml(xml1, xml2).ToString();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to merge XML elements of type {typeof(T).Name}", ex);
        }
    }

    /// <summary>
    /// Merges two XML documents with the second document's elements taking precedence when there are conflicts
    /// </summary>
    private static XDocument MergeXml(XDocument xd1, XDocument xd2)
    {
        if (xd1 == null || xd1.Root == null) return xd2;
        if (xd2 == null || xd2.Root == null) return xd1;

        // Create a new document with the root from the second document
        var result = new XDocument(
            new XElement(xd2.Root.Name,
                // Merge attributes, with the second document taking precedence
                xd1.Root.Attributes()
                    .Concat(xd2.Root.Attributes())
                    .GroupBy(g => g.Name)
                    .Select(g => g.Last()),
                    
                // Merge elements, with the second document taking precedence for elements with the same name
                MergeElements(xd1.Root.Elements(), xd2.Root.Elements())
            )
        );

        return result;
    }

    /// <summary>
    /// Merges two collections of XML elements with the second collection taking precedence in case of conflicts
    /// </summary>
    private static IEnumerable<XElement> MergeElements(IEnumerable<XElement> elements1, IEnumerable<XElement> elements2)
    {
        // Group elements by name
        var elementsByName = elements1
            .Concat(elements2)
            .GroupBy(e => e.Name)
            .ToList();

        foreach (var group in elementsByName)
        {
            var elementsInGroup = group.ToList();
            if (elementsInGroup.Count == 1)
            {
                // If there's only one element with this name, just return it
                yield return elementsInGroup[0];
            }
            else
            {
                // If there are multiple elements with the same name,
                // take the one from elements2 (which is later in the concatenation)
                var element1 = elementsInGroup.First();
                var element2 = elementsInGroup.Last();

                var mergedElement = new XElement(element2.Name,
                    // Merge attributes, with element2 taking precedence
                    element1.Attributes()
                        .Concat(element2.Attributes())
                        .GroupBy(a => a.Name)
                        .Select(g => g.Last()),
                        
                    // Merge child elements recursively
                    MergeElements(element1.Elements(), element2.Elements())
                );

                yield return mergedElement;
            }
        }
    }
}