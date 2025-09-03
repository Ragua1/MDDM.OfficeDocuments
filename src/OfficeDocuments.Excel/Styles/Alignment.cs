using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Extensions;

namespace OfficeDocuments.Excel.Styles;

/// <summary>
/// Class of Alignment
/// </summary>
public class Alignment
{
    /// <summary>
    /// Instance of Alignment element
    /// </summary>
    public DocumentFormat.OpenXml.Spreadsheet.Alignment Element { get; }
    /// <summary>
    /// Sets the horizontal.
    /// </summary>
    public HorizontalAlignmentValues Horizontal
    {
        set { Element.Horizontal = GetHorizontalAlignmentValues(value); }
    }
    /// <summary>
    /// Sets the vertical.
    /// </summary>
    public VerticalAlignmentValues Vertical
    {
        set { Element.Vertical = GetVerticalAlignmentValues(value); }
    }
    /// <summary>
    /// Sets a value indicating whether [justify last line].
    /// </summary>
    public bool JustifyLastLine
    {
        set { Element.JustifyLastLine = value; }
    }
    /// <summary>
    /// Sets a value indicating whether [wrap text].
    /// </summary>
    public bool WrapText
    {
        set { Element.WrapText = value; }
    }
    /// <summary>
    /// Sets a value indicating whether [shrink to fit].
    /// </summary>
    public bool ShrinkToFit
    {
        set { Element.ShrinkToFit = value; }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Alignment"/> class.
    /// </summary>
    /// <param name="alignment">Spreadsheet alignment</param>
    public Alignment(DocumentFormat.OpenXml.Spreadsheet.Alignment? alignment = null)
    {
        Element = alignment ?? new DocumentFormat.OpenXml.Spreadsheet.Alignment();
    }

    /// <summary>
    /// Compare content with 'alignment'
    /// </summary>
    /// <param name="alignment">Spreadsheet alignment for compare</param>
    public bool IsContentSame(DocumentFormat.OpenXml.Spreadsheet.Alignment alignment)
    {
        return alignment.OuterXml.CompareXml(Element.OuterXml);
    }
        
    private static DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues GetHorizontalAlignmentValues(HorizontalAlignmentValues horizontalAlignmentValues)
    {
        return horizontalAlignmentValues switch
        {
            HorizontalAlignmentValues.General => DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General,
            HorizontalAlignmentValues.Left => DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left,
            HorizontalAlignmentValues.Center => DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
            HorizontalAlignmentValues.Right => DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right,
            HorizontalAlignmentValues.Fill => DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Fill,
            HorizontalAlignmentValues.Justify => DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Justify,
            HorizontalAlignmentValues.CenterContinuous => DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.CenterContinuous,
            HorizontalAlignmentValues.Distributed => DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Distributed,
            _ => DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General
        };
    }
    private static DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues GetVerticalAlignmentValues(VerticalAlignmentValues verticalAlignmentValues)
    {
        return verticalAlignmentValues switch
        {
            VerticalAlignmentValues.Top => DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top,
            VerticalAlignmentValues.Center => DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center,
            VerticalAlignmentValues.Bottom => DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Bottom,
            VerticalAlignmentValues.Justify => DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Justify,
            VerticalAlignmentValues.Distributed => DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Distributed,
            _ => DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
        };
    }
}