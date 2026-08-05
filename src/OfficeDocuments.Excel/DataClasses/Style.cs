using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Styles;

namespace OfficeDocuments.Excel.DataClasses;

internal class Style : IStyle
{
    public uint StyleIndex { get; }

    internal DocumentFormat.OpenXml.Spreadsheet.Stylesheet StylesheetInternal { get; }
    internal DocumentFormat.OpenXml.Spreadsheet.CellFormat ElementInternal { get; }

#pragma warning disable CS0618
    DocumentFormat.OpenXml.Spreadsheet.Stylesheet IStyle.Stylesheet => StylesheetInternal;
    DocumentFormat.OpenXml.Spreadsheet.CellFormat IStyle.Element => ElementInternal;
#pragma warning restore CS0618

    public int FontId => Convert.ToInt32(ElementInternal.FontId?.Value ?? 0U);
    public int FillId => Convert.ToInt32(ElementInternal.FillId?.Value ?? 0U);
    public int BorderId => Convert.ToInt32(ElementInternal.BorderId?.Value ?? 0U);
    public int NumberFormatId => Convert.ToInt32(ElementInternal.NumberFormatId?.Value ?? 0);

    internal Style(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, Font? font = null, Fill? fill = null, Border? border = null, NumberingFormat? numberFormat = null)
        : this(stylesheet, GetFontId(stylesheet, font), GetFillId(stylesheet, fill), GetBorderId(stylesheet, border), GetNumberFormatId(stylesheet, numberFormat))
    { }
    internal Style(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, Font? font = null, Fill? fill = null, Border? border = null, NumberingFormat? numberFormat = null, Alignment? alignment = null)
        : this(stylesheet, GetFontId(stylesheet, font), GetFillId(stylesheet, fill), GetBorderId(stylesheet, border), GetNumberFormatId(stylesheet, numberFormat), alignment)
    { }
    internal Style(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, int? fontId = 0, int? fillId = 0, int? borderId = 0, int? numberFormatId = 0, Alignment? alignment = null)
    {
        StylesheetInternal = stylesheet;
        ElementInternal = new DocumentFormat.OpenXml.Spreadsheet.CellFormat
        {
            FormatId = Convert.ToUInt32(0),
            FontId = Convert.ToUInt32(fontId),
            FillId = Convert.ToUInt32(fillId),
            BorderId = Convert.ToUInt32(borderId)
        };

        if (numberFormatId >= 0)
        {
            ElementInternal.NumberFormatId = Convert.ToUInt32(numberFormatId);
        }

        if (alignment != null)
        {
            ElementInternal.Alignment = (DocumentFormat.OpenXml.Spreadsheet.Alignment)alignment.Element.CloneNode(true);
        }

        StyleIndex = GetStyleIndex(stylesheet, ElementInternal);
    }
    internal Style(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, uint styleIndex)
    {
        StylesheetInternal = stylesheet;

        var cfs = stylesheet.CellFormats ?? throw new InvalidOperationException("The stylesheet does not contain cell formats.");
        ElementInternal = GetElementAt<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(
            cfs,
            Convert.ToInt32(styleIndex),
            "cell format");

        StyleIndex = styleIndex;
    }

    public IStyle CreateMergedStyle(IStyle? style)
    {
        int fontId = FontId, fillId = FillId, borderId = BorderId, numberFormatId = NumberFormatId;
        var alignment = ElementInternal.Alignment != null ? new Alignment(ElementInternal.Alignment) : null;
        if (style == null)
        {
            return this;// new Style(this.Stylesheet, fontId, fillId, borderId, numberFormatId, alignment);
        }

        var sourceStylesheet = GetStylesheet(style);
        var sourceElement = GetElement(style);

        // Ids only mean the same thing inside one stylesheet. When merging across workbooks,
        // "same id" says nothing about "same content", so the skip-if-equal shortcut must not
        // apply — otherwise a facet is silently dropped whenever the two indexes coincide.
        var sameStylesheet = ReferenceEquals(sourceStylesheet, StylesheetInternal);
        var targetFonts = StylesheetInternal.Fonts;
        var sourceFonts = sourceStylesheet.Fonts;
        var targetFills = StylesheetInternal.Fills;
        var sourceFills = sourceStylesheet.Fills;
        var targetBorders = StylesheetInternal.Borders;
        var sourceBorders = sourceStylesheet.Borders;

        if (style.FontId > 0 && (!sameStylesheet || fontId != style.FontId))
        {
            var font1 = GetElementAt<DocumentFormat.OpenXml.Spreadsheet.Font>(
                targetFonts ?? throw new InvalidOperationException("The stylesheet does not contain fonts."),
                FontId,
                "font");
            var font2 = GetElementAt<DocumentFormat.OpenXml.Spreadsheet.Font>(
                sourceFonts ?? throw new InvalidOperationException("The source stylesheet does not contain fonts."),
                style.FontId,
                "font");
            var font = Utils.MergeFonts(font1, font2);
            fontId = GetFontId(StylesheetInternal, font);
        }

        if (style.FillId > 0 && (!sameStylesheet || fillId != style.FillId))
        {
            var fill1 = GetElementAt<DocumentFormat.OpenXml.Spreadsheet.Fill>(
                targetFills ?? throw new InvalidOperationException("The stylesheet does not contain fills."),
                FillId,
                "fill");
            var fill2 = GetElementAt<DocumentFormat.OpenXml.Spreadsheet.Fill>(
                sourceFills ?? throw new InvalidOperationException("The source stylesheet does not contain fills."),
                style.FillId,
                "fill");
            var fill = Utils.MergeFills(fill1, fill2);
            fillId = GetFillId(StylesheetInternal, fill);
        }

        if (style.BorderId > 0 && (!sameStylesheet || borderId != style.BorderId))
        {
            var border1 = GetElementAt<DocumentFormat.OpenXml.Spreadsheet.Border>(
                targetBorders ?? throw new InvalidOperationException("The stylesheet does not contain borders."),
                BorderId,
                "border");
            var border2 = GetElementAt<DocumentFormat.OpenXml.Spreadsheet.Border>(
                sourceBorders ?? throw new InvalidOperationException("The source stylesheet does not contain borders."),
                style.BorderId,
                "border");
            var border = Utils.MergeBorders(border1, border2);
            borderId = GetBorderId(StylesheetInternal, border);
        }

        if (style.NumberFormatId > 0 && (!sameStylesheet || numberFormatId != style.NumberFormatId))
        {
            if (style.NumberFormatId < NumberingFormat.FirstCustomNumberFormatId)
            {
                numberFormatId = style.NumberFormatId;
            }
            else
            {
                var sourceNumberingFormats = sourceStylesheet.NumberingFormats
                    ?? throw new InvalidOperationException("The source stylesheet does not contain numbering formats.");
                var sourceNumberFormat = FindNumberingFormatById(sourceNumberingFormats, style.NumberFormatId)
                    ?? throw new InvalidOperationException("The source stylesheet does not contain the expected number format.");

                numberFormatId = GetNumberFormatId(
                    StylesheetInternal,
                    new NumberingFormat((DocumentFormat.OpenXml.Spreadsheet.NumberingFormat)sourceNumberFormat.CloneNode(true))
                );
            }
        }

        if (HasAlignmentContent(sourceElement.Alignment))
        {
            if (!HasAlignmentContent(ElementInternal.Alignment))
            {
                alignment = new Alignment(sourceElement.Alignment); // Alignment cannot be merged
            }
        }

        return new Style(StylesheetInternal, fontId, fillId, borderId, numberFormatId, alignment);
    }

    private static int GetFontId(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, Font? font)
    {
        var fontId = 0;
        if (font?.Element != null)
        {
            var fonts = stylesheet.Fonts ?? (stylesheet.Fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts());
            fontId = FindElementIndex<DocumentFormat.OpenXml.Spreadsheet.Font>(fonts, font.IsContentSame);

            if (fontId < 0) // not found; a match at index 0 is the default entry and must be reused
            {
                fonts.Append(font.Element);
                fontId = fonts.ChildElements.Count - 1;
            }
        }
        return fontId;
    }

    private static int GetFillId(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, Fill? fill)
    {
        var fillId = 0;
        if (fill?.Element != null)
        {
            var fills = stylesheet.Fills ?? (stylesheet.Fills = new DocumentFormat.OpenXml.Spreadsheet.Fills());
            fillId = FindElementIndex<DocumentFormat.OpenXml.Spreadsheet.Fill>(fills, fill.IsContentSame);

            if (fillId < 0) // not found; a match at index 0 is the default entry and must be reused
            {
                fills.Append(fill.Element);
                fillId = fills.ChildElements.Count - 1;
            }
        }
        return fillId;
    }

    private static int GetBorderId(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, Border? border)
    {
        var borderId = 0;
        if (border?.Element != null)
        {
            var borders = stylesheet.Borders ?? (stylesheet.Borders = new DocumentFormat.OpenXml.Spreadsheet.Borders());
            borderId = FindElementIndex<DocumentFormat.OpenXml.Spreadsheet.Border>(borders, border.IsContentSame);

            if (borderId < 0) // not found; a match at index 0 is the default entry and must be reused
            {
                borders.Append(border.Element);
                borderId = borders.ChildElements.Count - 1;
            }
        }
        return borderId;
    }

    private static int GetNumberFormatId(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, NumberingFormat? numberFormat)
    {
        if (numberFormat?.Element == null)
        {
            return 0;
        }

        var formatCode = numberFormat.Element.FormatCode?.Value;
        if (NumberingFormat.TryGetBuiltInId(formatCode, out var builtInNumberFormatId))
        {
            return Convert.ToInt32(builtInNumberFormatId);
        }

        var numberingFormats = stylesheet.NumberingFormats ?? (stylesheet.NumberingFormats = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormats());
        var numFormat = FindElement(numberingFormats, numberFormat.IsContentSame);

        if (numFormat == null)
        {
            var nextNumberFormatId = NumberingFormat.GetNextCustomId(numberingFormats);
            var newNumberFormat = (DocumentFormat.OpenXml.Spreadsheet.NumberingFormat)numberFormat.Element.CloneNode(true);
            newNumberFormat.NumberFormatId = nextNumberFormatId;
            numberingFormats.Append(newNumberFormat);
            numberingFormats.Count = Convert.ToUInt32(numberingFormats.Count());

            return Convert.ToInt32(nextNumberFormatId);
        }

        return Convert.ToInt32(numFormat.NumberFormatId?.Value ?? 0U);
    }

    private static uint GetStyleIndex(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, DocumentFormat.OpenXml.Spreadsheet.CellFormat element)
    {
        var cfs = stylesheet.CellFormats ?? (stylesheet.CellFormats = new DocumentFormat.OpenXml.Spreadsheet.CellFormats());
        for (var i = 0; i < cfs.ChildElements.Count; i++)
        {
            if (cfs.ChildElements[i] is DocumentFormat.OpenXml.Spreadsheet.CellFormat existingElement && Equals(element, existingElement))
            {
                return Convert.ToUInt32(i);
            }
        }

        cfs.Append(element);
        cfs.Count = Convert.ToUInt32(cfs.Count());
        return (uint)(cfs.Count() - 1);
    }

    private static bool Equals(DocumentFormat.OpenXml.Spreadsheet.CellFormat style1, DocumentFormat.OpenXml.Spreadsheet.CellFormat style2)
    {
        var style1FontId = style1.FontId?.Value ?? 0U;
        var style2FontId = style2.FontId?.Value ?? 0U;
        var style1FillId = style1.FillId?.Value ?? 0U;
        var style2FillId = style2.FillId?.Value ?? 0U;
        var style1BorderId = style1.BorderId?.Value ?? 0U;
        var style2BorderId = style2.BorderId?.Value ?? 0U;

        var res = style1FontId == style2FontId
                  && style1FillId == style2FillId
                  && style1BorderId == style2BorderId;

        var style1NumberFormatId = style1.NumberFormatId?.Value;
        var style2NumberFormatId = style2.NumberFormatId?.Value;

        if (style1NumberFormatId.HasValue == style2NumberFormatId.HasValue)
        {
            if (style1NumberFormatId.HasValue)
            {
                res &= style1NumberFormatId.Value == style2NumberFormatId;
            }
        }
        else
        {
            return false;
        }

        if (style1.Alignment != null && style2.Alignment != null)
        {
            res &= Utils.OpenXmlElementsEqual(style1.Alignment, style2.Alignment);
        }
        else
        {
            if (style1.Alignment == null && style2.Alignment != null
                || style1.Alignment != null && style2.Alignment == null)
            {
                return false;
            }
        }


        return res;
    }

    private static DocumentFormat.OpenXml.Spreadsheet.Stylesheet GetStylesheet(IStyle style)
    {
        if (style is Style concreteStyle)
        {
            return concreteStyle.StylesheetInternal;
        }

#pragma warning disable CS0618
        return style.Stylesheet;
#pragma warning restore CS0618
    }

    private static DocumentFormat.OpenXml.Spreadsheet.CellFormat GetElement(IStyle style)
    {
        if (style is Style concreteStyle)
        {
            return concreteStyle.ElementInternal;
        }

#pragma warning disable CS0618
        return style.Element;
#pragma warning restore CS0618
    }

    private static TElement GetElementAt<TElement>(DocumentFormat.OpenXml.OpenXmlCompositeElement collection, int index, string elementName)
        where TElement : DocumentFormat.OpenXml.OpenXmlElement
    {
        if (index < 0 || index >= collection.ChildElements.Count || collection.ChildElements[index] is not TElement element)
        {
            throw new InvalidOperationException($"The stylesheet does not contain the expected {elementName}.");
        }

        return element;
    }

    private static int FindElementIndex<TElement>(DocumentFormat.OpenXml.OpenXmlCompositeElement collection, Func<TElement, bool> predicate)
        where TElement : DocumentFormat.OpenXml.OpenXmlElement
    {
        for (var i = 0; i < collection.ChildElements.Count; i++)
        {
            if (collection.ChildElements[i] is TElement element && predicate(element))
            {
                return i;
            }
        }

        return -1;
    }

    private static DocumentFormat.OpenXml.Spreadsheet.NumberingFormat? FindElement(
        DocumentFormat.OpenXml.OpenXmlCompositeElement collection,
        Func<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat, bool> predicate)
    {
        for (var i = 0; i < collection.ChildElements.Count; i++)
        {
            if (collection.ChildElements[i] is DocumentFormat.OpenXml.Spreadsheet.NumberingFormat element && predicate(element))
            {
                return element;
            }
        }

        return null;
    }

    private static DocumentFormat.OpenXml.Spreadsheet.NumberingFormat? FindNumberingFormatById(
        DocumentFormat.OpenXml.Spreadsheet.NumberingFormats numberingFormats,
        int numberFormatId)
    {
        for (var i = 0; i < numberingFormats.ChildElements.Count; i++)
        {
            if (numberingFormats.ChildElements[i] is DocumentFormat.OpenXml.Spreadsheet.NumberingFormat format
                && Convert.ToInt32(format.NumberFormatId?.Value ?? 0U) == numberFormatId)
            {
                return format;
            }
        }

        return null;
    }

    private static bool HasAlignmentContent(DocumentFormat.OpenXml.Spreadsheet.Alignment? alignment)
    {
        return alignment != null && (alignment.HasAttributes || alignment.ChildElements.Count > 0);
    }
}