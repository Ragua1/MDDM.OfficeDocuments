using OfficeDocuments.Word.Enums;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Table-level formatting: width, alignment, borders, and cell padding.
/// </summary>
/// <remarks>
/// As with the other format records, <see langword="null"/> means "leave this alone", so a format can
/// carry one change without restating the rest. Measurements are in points except
/// <see cref="WidthPercent"/>.
/// </remarks>
public sealed record TableFormat
{
    private readonly double? _widthPercent;
    private readonly double? _borderWidth;
    private readonly double? _cellPadding;
    private readonly string? _borderColor;
    private readonly string? _styleId;

    /// <summary>
    /// Table width as a percentage of the available text width, from 0 to 100.
    /// </summary>
    /// <remarks>
    /// A percentage rather than an absolute width because it is what survives a change of page size
    /// or margins. Leave it unset to let Word size the table from its content.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">The value is outside 0 to 100.</exception>
    public double? WidthPercent
    {
        get => _widthPercent;
        init
        {
            if (value is not null && (double.IsNaN(value.Value) || value.Value < 0d || value.Value > 100d))
            {
                throw new ArgumentOutOfRangeException(nameof(WidthPercent), value, "Table width must be between 0 and 100 percent.");
            }

            _widthPercent = value;
        }
    }

    /// <summary>
    /// Horizontal placement of the table between the margins.
    /// </summary>
    public TableAlignment? Alignment { get; init; }

    /// <summary>
    /// Which edges get borders. Use <see cref="TableBorders.None"/> to remove them.
    /// </summary>
    public TableBorders? Borders { get; init; }

    /// <summary>
    /// Border colour as 6 hex digits (<c>RRGGBB</c>), optionally <c>#</c>-prefixed, or <c>auto</c>.
    /// </summary>
    /// <exception cref="ArgumentException">The value is not a colour Word can store.</exception>
    public string? BorderColor
    {
        get => _borderColor;
        init => _borderColor = value is null ? null : HexColor.Normalize(value, nameof(BorderColor));
    }

    /// <summary>
    /// Border line width in points. Stored in eighths of a point, so 0.125 pt is the finest step.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative or not finite.</exception>
    public double? BorderWidth
    {
        get => _borderWidth;
        init => _borderWidth = value is null ? null : Measure.ValidateLength(value.Value, nameof(BorderWidth));
    }

    /// <summary>
    /// Space between a cell's border and its content, in points, applied to every side.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative or not finite.</exception>
    public double? CellPadding
    {
        get => _cellPadding;
        init => _cellPadding = value is null ? null : Measure.ValidateLength(value.Value, nameof(CellPadding));
    }

    /// <summary>
    /// Identifier of a table style defined in the document, for a document created from a template.
    /// </summary>
    /// <remarks>
    /// This library defines no built-in table styles, so an identifier here is written through
    /// untouched and only renders if the document already defines it.
    /// </remarks>
    /// <exception cref="ArgumentException">The identifier is empty or whitespace.</exception>
    public string? StyleId
    {
        get => _styleId;
        init
        {
            if (value is not null)
            {
                ArgumentException.ThrowIfNullOrWhiteSpace(value, nameof(StyleId));
            }

            _styleId = value;
        }
    }

    /// <summary>
    /// <see langword="true"/> when no property is set, so applying it would write nothing.
    /// </summary>
    public bool IsEmpty =>
        WidthPercent is null
        && Alignment is null
        && Borders is null
        && BorderColor is null
        && BorderWidth is null
        && CellPadding is null
        && StyleId is null;

    /// <summary>
    /// Layers <paramref name="overrides"/> on top of this format: every property the argument sets
    /// wins, and the ones it leaves unset keep this format's value.
    /// </summary>
    /// <param name="overrides">Format whose set properties take precedence. May be <see langword="null"/>.</param>
    /// <returns>The combined format. Neither input is modified.</returns>
    public TableFormat Merge(TableFormat? overrides)
    {
        if (overrides is null || overrides.IsEmpty)
        {
            return this;
        }

        return new TableFormat
        {
            WidthPercent = overrides.WidthPercent ?? WidthPercent,
            Alignment = overrides.Alignment ?? Alignment,
            Borders = overrides.Borders ?? Borders,
            BorderColor = overrides.BorderColor ?? BorderColor,
            BorderWidth = overrides.BorderWidth ?? BorderWidth,
            CellPadding = overrides.CellPadding ?? CellPadding,
            StyleId = overrides.StyleId ?? StyleId,
        };
    }
}
