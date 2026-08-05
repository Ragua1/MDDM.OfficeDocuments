using System.Collections.Generic;
using System.Linq;
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
    /// Normalizes a user-supplied color string into the 8-digit uppercase ARGB form that the
    /// OOXML <c>rgb</c> attribute requires (it is typed as hexBinary, so a leading '#' or a
    /// 6-digit RGB value produces a file Excel refuses to open).
    /// </summary>
    /// <param name="value">A 6- or 8-digit hex color, optionally prefixed with '#'.</param>
    /// <param name="parameterName">Name reported when <paramref name="value"/> is rejected.</param>
    internal static string NormalizeArgbHex(string value, string parameterName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(value, parameterName);

        var hex = value.AsSpan().Trim();
        if (hex.Length > 0 && hex[0] == '#')
        {
            hex = hex[1..];
        }

        if (hex.Length is not (6 or 8) || !IsHexDigits(hex))
        {
            throw new ArgumentException(
                $"'{value}' is not a valid color. Expected 6 (RGB) or 8 (ARGB) hexadecimal digits, optionally prefixed with '#'.",
                parameterName);
        }

        var normalized = hex.ToString().ToUpperInvariant();

        return hex.Length == 6 ? $"FF{normalized}" : normalized;
    }

    private static bool IsHexDigits(ReadOnlySpan<char> value)
    {
        foreach (var character in value)
        {
            if (!char.IsAsciiHexDigit(character))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Create new font by merging two fonts
    /// </summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    public static Font MergeFonts(DocumentFormat.OpenXml.Spreadsheet.Font font1, DocumentFormat.OpenXml.Spreadsheet.Font font2)
    {
        if (font1 == null) throw new ArgumentNullException(nameof(font1));
        if (font2 == null) throw new ArgumentNullException(nameof(font2));

        return new Font(MergeElements(font1, font2));
    }

    /// <summary>
    /// Create new fill by merging two fills
    /// </summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    public static Fill MergeFills(DocumentFormat.OpenXml.Spreadsheet.Fill fill1, DocumentFormat.OpenXml.Spreadsheet.Fill fill2)
    {
        if (fill1 == null) throw new ArgumentNullException(nameof(fill1));
        if (fill2 == null) throw new ArgumentNullException(nameof(fill2));

        return new Fill(MergeElements(fill1, fill2));
    }

    /// <summary>
    /// Create new border by merging two borders
    /// </summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    public static Border MergeBorders(DocumentFormat.OpenXml.Spreadsheet.Border border1, DocumentFormat.OpenXml.Spreadsheet.Border border2)
    {
        if (border1 == null) throw new ArgumentNullException(nameof(border1));
        if (border2 == null) throw new ArgumentNullException(nameof(border2));

        return new Border(MergeElements(border1, border2));
    }

    internal static bool OpenXmlElementsEqual(OpenXmlElement? element1, OpenXmlElement? element2)
    {
        if (ReferenceEquals(element1, element2))
        {
            return true;
        }

        if (element1 == null || element2 == null)
        {
            return false;
        }

        if (element1.GetType() != element2.GetType())
        {
            return false;
        }

        if (!HaveSameAttributes(element1, element2))
        {
            return false;
        }

        if (!string.Equals(element1.InnerText, element2.InnerText, StringComparison.Ordinal))
        {
            return false;
        }

        if (element1.ChildElements.Count != element2.ChildElements.Count)
        {
            return false;
        }

        var unmatchedChildren = element2.ChildElements.Select(child => child).ToList();
        foreach (var child1 in element1.ChildElements)
        {
            var matchIndex = unmatchedChildren.FindIndex(child2 => HaveSameIdentity(child1, child2) && OpenXmlElementsEqual(child1, child2));
            if (matchIndex < 0)
            {
                return false;
            }

            unmatchedChildren.RemoveAt(matchIndex);
        }

        return unmatchedChildren.Count == 0;
    }

    private static T MergeElements<T>(T element1, T element2)
        where T : OpenXmlElement
    {
        if (element1.GetType() != element2.GetType())
        {
            throw new InvalidOperationException($"Cannot merge different OpenXml element types: {element1.GetType().Name} and {element2.GetType().Name}.");
        }

        if (element1 is not OpenXmlCompositeElement composite1 || element2 is not OpenXmlCompositeElement composite2)
        {
            return (T)element2.CloneNode(true);
        }

        var merged = (T)element1.CloneNode(true);
        ApplyAttributes(merged, element2);

        if (merged is not OpenXmlCompositeElement mergedComposite)
        {
            return merged;
        }

        foreach (var overrideChild in composite2.ChildElements)
        {
            var existingChild = mergedComposite.ChildElements.FirstOrDefault(candidate => HaveSameIdentity(candidate, overrideChild));
            var mergedChild = existingChild != null
                ? MergeElements(existingChild, overrideChild)
                : overrideChild.CloneNode(true);

            if (existingChild == null)
            {
                mergedComposite.Append(mergedChild);
            }
            else
            {
                mergedComposite.ReplaceChild(mergedChild, existingChild);
            }
        }

        ApplySchemaChildOrder(mergedComposite);

        return merged;
    }

    /// <summary>
    /// Style child elements are declared as an xsd:sequence, so a merged-in child appended at the
    /// end produces a schema-invalid file whenever it belongs before a child the base already had
    /// (for example a bold merged onto a font that only carried a size).
    /// </summary>
    private static readonly Dictionary<string, string[]> ChildOrderByParent = new(StringComparer.Ordinal)
    {
        // CT_Font (ECMA-376 Part 1, 18.8.22)
        ["font"] =
        [
            "b", "i", "strike", "condense", "extend", "outline", "shadow", "u", "vertAlign",
            "sz", "color", "name", "family", "charset", "scheme"
        ],
        // CT_Border (18.8.4)
        ["border"] = ["left", "right", "top", "bottom", "diagonal", "vertical", "horizontal"],
        // CT_PatternFill (18.8.32)
        ["patternFill"] = ["fgColor", "bgColor"]
    };

    private static void ApplySchemaChildOrder(OpenXmlCompositeElement element)
    {
        if (!ChildOrderByParent.TryGetValue(element.LocalName, out var order))
        {
            return;
        }

        var children = element.ChildElements.ToList();
        if (children.Count < 2)
        {
            return;
        }

        // OrderBy is stable, so children outside the known sequence keep their relative order.
        var sorted = children.OrderBy(child => PositionOf(child.LocalName, order)).ToList();
        if (sorted.SequenceEqual(children))
        {
            return;
        }

        element.RemoveAllChildren();
        foreach (var child in sorted)
        {
            element.AppendChild(child);
        }
    }

    private static int PositionOf(string localName, string[] order)
    {
        var index = Array.IndexOf(order, localName);

        return index < 0 ? int.MaxValue : index;
    }

    private static void ApplyAttributes(OpenXmlElement target, OpenXmlElement source)
    {
        foreach (var attribute in source.GetAttributes())
        {
            target.SetAttribute(attribute);
        }
    }

    private static bool HaveSameAttributes(OpenXmlElement element1, OpenXmlElement element2)
    {
        var attributes1 = element1.GetAttributes()
            .OrderBy(attribute => attribute.NamespaceUri, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.LocalName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Prefix, StringComparer.Ordinal)
            .ToList();
        var attributes2 = element2.GetAttributes()
            .OrderBy(attribute => attribute.NamespaceUri, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.LocalName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Prefix, StringComparer.Ordinal)
            .ToList();

        if (attributes1.Count != attributes2.Count)
        {
            return false;
        }

        for (var index = 0; index < attributes1.Count; index++)
        {
            var left = attributes1[index];
            var right = attributes2[index];
            if (!string.Equals(left.NamespaceUri, right.NamespaceUri, StringComparison.Ordinal)
                || !string.Equals(left.LocalName, right.LocalName, StringComparison.Ordinal)
                || !string.Equals(left.Prefix, right.Prefix, StringComparison.Ordinal)
                || !string.Equals(left.Value, right.Value, StringComparison.Ordinal))
            {
                return false;
            }
        }

        return true;
    }

    private static bool HaveSameIdentity(OpenXmlElement element1, OpenXmlElement element2)
    {
        return element1.GetType() == element2.GetType()
               && string.Equals(element1.LocalName, element2.LocalName, StringComparison.Ordinal)
               && string.Equals(element1.NamespaceUri, element2.NamespaceUri, StringComparison.Ordinal);
    }
}