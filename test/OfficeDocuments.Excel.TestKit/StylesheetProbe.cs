using OfficeDocuments.Excel.Interfaces;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.TestKit;

/// <summary>
/// Reads the entries a style actually points at, so tests can assert on the produced
/// <c>styles.xml</c> rather than on the allocated index.
/// </summary>
/// <remarks>
/// An id assertion only proves that <em>some</em> entry was allocated — it cannot tell a blue font
/// from a red one. This type is also where the obsolete raw-stylesheet access is confined, so the
/// test projects do not each carry their own <c>CS0618</c> suppression.
/// </remarks>
public static class StylesheetProbe
{
#pragma warning disable CS0618 // the raw stylesheet is the only way to read rendered style state
    private static SpreadsheetLib.Stylesheet Sheet(IStyle style) => style.Stylesheet;

    /// <summary>
    /// The <c>cellXfs</c> entry the style resolves to.
    /// </summary>
    public static SpreadsheetLib.CellFormat CellFormat(IStyle style) => style.Element;

    /// <summary>
    /// Whether both styles were allocated in the same stylesheet — the check that a merge across
    /// workbooks landed in the target workbook rather than the source one.
    /// </summary>
    public static bool ShareStylesheet(IStyle first, IStyle second) =>
        ReferenceEquals(first.Stylesheet, second.Stylesheet);
#pragma warning restore CS0618

    /// <summary>
    /// The <c>alignment</c> the style carries, or <see langword="null"/> when it sets none.
    /// </summary>
    /// <remarks>
    /// <see cref="IStyle"/> has no alignment accessor, so this is the only way to tell "no
    /// alignment" from "an alignment that happens to be the default".
    /// </remarks>
    public static SpreadsheetLib.Alignment? Alignment(IStyle style) => CellFormat(style).Alignment;

    /// <summary>The <c>font</c> entry the style points at.</summary>
    public static SpreadsheetLib.Font Font(IStyle style) =>
        ElementAt<SpreadsheetLib.Font>(Sheet(style).Fonts, style.FontId, "font");

    /// <summary>The <c>fill</c> entry the style points at.</summary>
    public static SpreadsheetLib.Fill Fill(IStyle style) =>
        ElementAt<SpreadsheetLib.Fill>(Sheet(style).Fills, style.FillId, "fill");

    /// <summary>The <c>border</c> entry the style points at.</summary>
    public static SpreadsheetLib.Border Border(IStyle style) =>
        ElementAt<SpreadsheetLib.Border>(Sheet(style).Borders, style.BorderId, "border");

    /// <summary>
    /// The custom <c>numFmt</c> the style points at, or <see langword="null"/> for a built-in id.
    /// </summary>
    public static SpreadsheetLib.NumberingFormat? NumberingFormat(IStyle style) =>
        Sheet(style).NumberingFormats?
            .Elements<SpreadsheetLib.NumberingFormat>()
            .FirstOrDefault(format => format.NumberFormatId?.Value == (uint)style.NumberFormatId);

    /// <summary>Number of <c>font</c> entries in the stylesheet, defaults included.</summary>
    public static int FontCount(IStyle style) => Count<SpreadsheetLib.Font>(Sheet(style).Fonts);

    /// <summary>Number of <c>fill</c> entries in the stylesheet, defaults included.</summary>
    public static int FillCount(IStyle style) => Count<SpreadsheetLib.Fill>(Sheet(style).Fills);

    /// <summary>Number of <c>border</c> entries in the stylesheet, defaults included.</summary>
    public static int BorderCount(IStyle style) => Count<SpreadsheetLib.Border>(Sheet(style).Borders);

    /// <summary>Number of <c>cellXfs</c> entries in the stylesheet, defaults included.</summary>
    public static int CellFormatCount(IStyle style) => Count<SpreadsheetLib.CellFormat>(Sheet(style).CellFormats);

    /// <summary>Number of custom <c>numFmt</c> entries in the stylesheet.</summary>
    public static int NumberingFormatCount(IStyle style) =>
        Count<SpreadsheetLib.NumberingFormat>(Sheet(style).NumberingFormats);

    private static int Count<TElement>(DocumentFormat.OpenXml.OpenXmlCompositeElement? collection)
        where TElement : DocumentFormat.OpenXml.OpenXmlElement =>
        collection?.Elements<TElement>().Count() ?? 0;

    private static TElement ElementAt<TElement>(DocumentFormat.OpenXml.OpenXmlCompositeElement? collection, int index, string elementName)
        where TElement : DocumentFormat.OpenXml.OpenXmlElement
    {
        if (collection is null)
        {
            throw new InvalidOperationException($"The stylesheet contains no {elementName} collection.");
        }

        var elements = collection.Elements<TElement>().ToList();
        if (index < 0 || index >= elements.Count)
        {
            throw new InvalidOperationException(
                $"The stylesheet has {elements.Count} {elementName} entries; index {index} is out of range.");
        }

        return elements[index];
    }
}
