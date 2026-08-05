using System.ComponentModel;
using OfficeDocuments.Excel.DataClasses;
using OfficeDocuments.Excel.Interfaces;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;
using Alignment = OfficeDocuments.Excel.Styles.Alignment;
using Border = OfficeDocuments.Excel.Styles.Border;
using Fill = OfficeDocuments.Excel.Styles.Fill;
using Font = OfficeDocuments.Excel.Styles.Font;
using NumberingFormat = OfficeDocuments.Excel.Styles.NumberingFormat;

namespace OfficeDocuments.Excel;

public partial class Spreadsheet
{
    public IStyle CreateStyle(Font? font = null, Fill? fill = null, Border? border = null, NumberingFormat? numberFormat = null, Alignment? alignment = null)
    {
        return new Style(StylesheetInternal, font, fill, border, numberFormat, alignment);
    }

    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This overload exposes raw OpenXml stylesheet plumbing. Prefer CreateStyle(...) without a Stylesheet parameter.")]
    public IStyle CreateStyle(SpreadsheetLib.Stylesheet stylesheet, Font? font = null, Fill? fill = null, Border? border = null, NumberingFormat? numberFormat = null, Alignment? alignment = null)
    {
        ArgumentNullException.ThrowIfNull(stylesheet);
        return new Style(stylesheet, font, fill, border, numberFormat, alignment);
    }

    internal uint GetOrCreateDifferentialFormat(IStyle style)
    {
        var differentialFormats = StylesheetInternal.DifferentialFormats ??= new SpreadsheetLib.DifferentialFormats();
        var differentialFormat = CreateDifferentialFormat(style);
        var existingFormats = differentialFormats.Elements<SpreadsheetLib.DifferentialFormat>().ToList();
        var existingIndex = existingFormats.FindIndex(existing => Utils.OpenXmlElementsEqual(existing, differentialFormat));
        if (existingIndex >= 0)
        {
            return Convert.ToUInt32(existingIndex);
        }

        differentialFormats.Append(differentialFormat);
        differentialFormats.Count = Convert.ToUInt32(differentialFormats.Count());
        return Convert.ToUInt32(existingFormats.Count);
    }

    private SpreadsheetLib.DifferentialFormat CreateDifferentialFormat(IStyle style)
    {
        var differentialFormat = new SpreadsheetLib.DifferentialFormat();

        if (style.FontId > 0)
        {
            var font = StylesheetInternal.Fonts?.Elements<SpreadsheetLib.Font>().ElementAt(style.FontId);
            if (font != null)
            {
                differentialFormat.Font = (SpreadsheetLib.Font)font.CloneNode(true);
            }
        }

        if (style.FillId > 0)
        {
            var fill = StylesheetInternal.Fills?.Elements<SpreadsheetLib.Fill>().ElementAt(style.FillId);
            if (fill != null)
            {
                differentialFormat.Fill = (SpreadsheetLib.Fill)fill.CloneNode(true);
            }
        }

        if (style.BorderId > 0)
        {
            var border = StylesheetInternal.Borders?.Elements<SpreadsheetLib.Border>().ElementAt(style.BorderId);
            if (border != null)
            {
                differentialFormat.Border = (SpreadsheetLib.Border)border.CloneNode(true);
            }
        }

        var styleElement = GetStyleElement(style);
        if (styleElement.Alignment != null)
        {
            differentialFormat.Alignment = (SpreadsheetLib.Alignment)styleElement.Alignment.CloneNode(true);
        }

        return differentialFormat;
    }

    private static SpreadsheetLib.CellFormat GetStyleElement(IStyle style)
    {
        if (style is Style concreteStyle)
        {
            return concreteStyle.ElementInternal;
        }

#pragma warning disable CS0618
        return style.Element;
#pragma warning restore CS0618
    }
}
