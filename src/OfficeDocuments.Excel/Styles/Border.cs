using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Extensions;

namespace OfficeDocuments.Excel.Styles;

/// <summary>
/// Class of Border
/// </summary>
public class Border
{
    /// <summary>
    /// Instance of Border element
    /// </summary>
    public DocumentFormat.OpenXml.Spreadsheet.Border Element { get; }
    /// <summary>
    /// Sets the left border
    /// </summary>
    public BorderStyleValues Left
    {
        set { Element.LeftBorder = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder { Style = GetBorderStyleValue(value) }; }
    }
    /// <summary>
    /// Sets the right border
    /// </summary>
    public BorderStyleValues Right
    {
        set { Element.RightBorder = new DocumentFormat.OpenXml.Spreadsheet.RightBorder { Style = GetBorderStyleValue(value) }; }
    }
    /// <summary>
    /// Sets the top border
    /// </summary>
    public BorderStyleValues Top
    {
        set { Element.TopBorder = new DocumentFormat.OpenXml.Spreadsheet.TopBorder { Style = GetBorderStyleValue(value) }; }
    }
    /// <summary>
    /// Sets the bottom border
    /// </summary>
    public BorderStyleValues Bottom
    {
        set { Element.BottomBorder = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder { Style = GetBorderStyleValue(value) }; }
    }
    /// <summary>
    /// Sets the set border style.
    /// </summary>
    public BorderStyleValues SetBorderStyle
    {
        set
        {
            var style = GetBorderStyleValue(value);
            Element.TopBorder = new DocumentFormat.OpenXml.Spreadsheet.TopBorder { Style = style };
            Element.RightBorder = new DocumentFormat.OpenXml.Spreadsheet.RightBorder { Style = style };
            Element.BottomBorder = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder { Style = style };
            Element.LeftBorder = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder { Style = style };
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Border"/> class.
    /// </summary>
    /// <param name="border">Spreadsheet border</param>
    public Border(DocumentFormat.OpenXml.Spreadsheet.Border border = null)
    {
        Element = border ?? new DocumentFormat.OpenXml.Spreadsheet.Border();
    }
    /// <summary>
    /// Initializes a new instance of the <see cref="Border"/> class.
    /// </summary>
    /// <param name="borderStyle">The border style.</param>
    public Border(BorderStyleValues borderStyle)
        : this()
    {
        SetBorderStyle = borderStyle;
    }

    /// <summary>
    /// Compare content with 'border'
    /// </summary>
    /// <param name="border">Spreadsheet border for compare</param>
    public bool IsContentSame(DocumentFormat.OpenXml.Spreadsheet.Border border)
    {
        return border.OuterXml.CompareXml(Element.OuterXml);
    }
        
    private static DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues GetBorderStyleValue(BorderStyleValues borderStyle)
    {
        return borderStyle switch
        {
            BorderStyleValues.None => DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.None,
            BorderStyleValues.Thin => DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin,
            BorderStyleValues.Medium => DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Medium,
            BorderStyleValues.Dashed => DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Dashed,
            BorderStyleValues.Dotted => DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Dotted,
            BorderStyleValues.Thick => DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick,
            BorderStyleValues.Double => DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Double,
            BorderStyleValues.Hair => DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Hair,
            BorderStyleValues.MediumDashed => DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.MediumDashed,
            BorderStyleValues.DashDot => DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.DashDot,
            _ => DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.None,

        };
    }
}