using System.ComponentModel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.Interfaces;

/// <summary>
/// Interface of style
/// </summary>
public interface IStyle
{
    /// <summary>
    /// Instance of OpenXml Stylesheet
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This property exposes raw OpenXml stylesheet plumbing. Use style metadata properties (FontId, FillId, BorderId, NumberFormatId) or CreateMergedStyle instead.")]
    Stylesheet Stylesheet { get; }
    /// <summary>
    /// Instance of style element
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This property exposes the raw OpenXml CellFormat. Use style metadata properties (FontId, FillId, BorderId, NumberFormatId) or CreateMergedStyle instead.")]
    CellFormat Element { get; }
    /// <summary>
    /// Style index
    /// </summary>
    uint StyleIndex { get; }

    /// <summary>
    /// Font id of style
    /// </summary>
    int FontId { get; }
    /// <summary>
    /// Fill id of style
    /// </summary>
    int FillId { get; }
    /// <summary>
    /// Border id of style
    /// </summary>
    int BorderId { get; }
    /// <summary>
    /// number format id of style
    /// </summary>
    int NumberFormatId { get; }

    /// <summary>
    /// Compare and merge style with 'style'
    /// </summary>
    /// <param name="style"></param>
    /// <returns>Merged style</returns>
    IStyle CreateMergedStyle(IStyle? style);
}