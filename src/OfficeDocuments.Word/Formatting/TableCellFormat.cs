using OfficeDocuments.Word.Enums;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Cell-level formatting: width, background, vertical alignment, and column spanning.
/// </summary>
public sealed record TableCellFormat
{
    private readonly double? _widthPercent;
    private readonly string? _backgroundColor;
    private readonly int? _columnSpan;

    /// <summary>
    /// Cell width as a percentage of the table width, from 0 to 100.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is outside 0 to 100.</exception>
    public double? WidthPercent
    {
        get => _widthPercent;
        init
        {
            if (value is not null && (double.IsNaN(value.Value) || value.Value < 0d || value.Value > 100d))
            {
                throw new ArgumentOutOfRangeException(nameof(WidthPercent), value, "Cell width must be between 0 and 100 percent.");
            }

            _widthPercent = value;
        }
    }

    /// <summary>
    /// Cell shading as 6 hex digits (<c>RRGGBB</c>), optionally <c>#</c>-prefixed, or <c>auto</c>.
    /// </summary>
    /// <exception cref="ArgumentException">The value is not a colour Word can store.</exception>
    public string? BackgroundColor
    {
        get => _backgroundColor;
        init => _backgroundColor = value is null ? null : HexColor.Normalize(value, nameof(BackgroundColor));
    }

    /// <summary>
    /// Vertical placement of the cell's content.
    /// </summary>
    public CellVerticalAlignment? VerticalAlignment { get; init; }

    /// <summary>
    /// Number of grid columns this cell occupies. 1 is a normal cell.
    /// </summary>
    /// <remarks>
    /// A spanned cell replaces the cells it covers rather than sitting alongside them, so a row with a
    /// cell of span 2 holds one fewer cell than the table has columns.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
    public int? ColumnSpan
    {
        get => _columnSpan;
        init
        {
            if (value is < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(ColumnSpan), value, "Column span must be at least 1.");
            }

            _columnSpan = value;
        }
    }

    /// <summary>
    /// <see langword="true"/> when no property is set, so applying it would write nothing.
    /// </summary>
    public bool IsEmpty =>
        WidthPercent is null
        && BackgroundColor is null
        && VerticalAlignment is null
        && ColumnSpan is null;

    /// <summary>
    /// Layers <paramref name="overrides"/> on top of this format: every property the argument sets
    /// wins, and the ones it leaves unset keep this format's value.
    /// </summary>
    /// <param name="overrides">Format whose set properties take precedence. May be <see langword="null"/>.</param>
    /// <returns>The combined format. Neither input is modified.</returns>
    public TableCellFormat Merge(TableCellFormat? overrides)
    {
        if (overrides is null || overrides.IsEmpty)
        {
            return this;
        }

        return new TableCellFormat
        {
            WidthPercent = overrides.WidthPercent ?? WidthPercent,
            BackgroundColor = overrides.BackgroundColor ?? BackgroundColor,
            VerticalAlignment = overrides.VerticalAlignment ?? VerticalAlignment,
            ColumnSpan = overrides.ColumnSpan ?? ColumnSpan,
        };
    }
}
