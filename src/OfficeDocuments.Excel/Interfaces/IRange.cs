using System.Collections.Generic;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Options;

namespace OfficeDocuments.Excel.Interfaces;

/// <summary>
/// Represents a rectangular worksheet area.
/// </summary>
public interface IRange : IBase
{
    /// <summary>
    /// First column of the range.
    /// </summary>
    uint FromColumn { get; }

    /// <summary>
    /// First row of the range.
    /// </summary>
    uint FromRow { get; }

    /// <summary>
    /// Last column of the range.
    /// </summary>
    uint ToColumn { get; }

    /// <summary>
    /// Last row of the range.
    /// </summary>
    uint ToRow { get; }

    /// <summary>
    /// Range reference in A1 notation.
    /// </summary>
    string Reference { get; }

    /// <summary>
    /// First cell reference in A1 notation.
    /// </summary>
    string StartReference { get; }

    /// <summary>
    /// Last cell reference in A1 notation.
    /// </summary>
    string EndReference { get; }

    /// <summary>
    /// Materialized rows in the range.
    /// </summary>
    IReadOnlyList<IRow> Rows { get; }

    /// <summary>
    /// Materialized cells in the range.
    /// </summary>
    IReadOnlyList<ICell> Cells { get; }

    /// <summary>
    /// Gets a cell in the range by absolute worksheet coordinates.
    /// </summary>
    ICell? GetCell(uint columnIndex, uint rowIndex);

    /// <summary>
    /// Gets a cell in the range by A1 reference.
    /// </summary>
    ICell? GetCell(string reference);

    /// <summary>
    /// Gets the range values row-by-row.
    /// Formula cells return the formula text.
    /// </summary>
    IReadOnlyList<IReadOnlyList<string?>> GetValues();

    /// <summary>
    /// Writes values to the range from the top-left corner.
    /// </summary>
    void SetValues(IEnumerable<IEnumerable<object?>> values);

    /// <summary>
    /// Applies a style to every cell in the range.
    /// </summary>
    void ApplyStyle(IStyle? style);

    /// <summary>
    /// Merges the range.
    /// </summary>
    void Merge();

    /// <summary>
    /// Applies an autofilter to the range.
    /// </summary>
    void ApplyAutoFilter();

    /// <summary>
    /// Sorts the range by a relative column index.
    /// </summary>
    void SortByColumn(uint relativeColumnIndex, SortDirection direction = SortDirection.Ascending, bool hasHeader = false);

    /// <summary>
    /// Applies a data validation rule to the range.
    /// </summary>
    void AddValidation(DataValidationOptions options);

    /// <summary>
    /// Applies a conditional formatting rule to the range.
    /// </summary>
    void AddConditionalFormatting(ConditionalFormattingOptions options);
}
