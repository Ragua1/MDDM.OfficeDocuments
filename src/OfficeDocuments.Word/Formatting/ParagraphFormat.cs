using OfficeDocuments.Word.Enums;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Paragraph-level formatting: alignment, spacing, indentation, and the named style to apply.
/// </summary>
/// <remarks>
/// <para>
/// As with <see cref="TextFormat"/>, <see langword="null"/> means "leave this alone", so a format
/// can carry a single change without restating everything else.
/// </para>
/// <para>
/// All measurements are in points. WordprocessingML stores them in twentieths of a point, and line
/// spacing in units of 240 per line; those conversions happen inside the library.
/// </para>
/// </remarks>
public sealed record ParagraphFormat
{
    /// <summary>
    /// Deepest list nesting level WordprocessingML numbering definitions carry.
    /// </summary>
    internal const int MaxListLevel = 8;

    private readonly int? _listLevel;
    private readonly double? _spacingBefore;
    private readonly double? _spacingAfter;
    private readonly double? _lineSpacing;
    private readonly double? _indentLeft;
    private readonly double? _indentRight;
    private readonly double? _indentFirstLine;
    private readonly string? _styleId;

    /// <summary>
    /// Horizontal alignment of the paragraph's text.
    /// </summary>
    public ParagraphAlignment? Alignment { get; init; }

    /// <summary>
    /// Space above the paragraph, in points.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative or not finite.</exception>
    public double? SpacingBefore
    {
        get => _spacingBefore;
        init => _spacingBefore = value is null ? null : Measure.ValidateLength(value.Value, nameof(SpacingBefore));
    }

    /// <summary>
    /// Space below the paragraph, in points.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative or not finite.</exception>
    public double? SpacingAfter
    {
        get => _spacingAfter;
        init => _spacingAfter = value is null ? null : Measure.ValidateLength(value.Value, nameof(SpacingAfter));
    }

    /// <summary>
    /// Line spacing as a multiple of single spacing: <c>1.0</c> is single, <c>1.5</c> is one-and-a-half,
    /// <c>2.0</c> is double.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not a positive finite number.</exception>
    public double? LineSpacing
    {
        get => _lineSpacing;
        init
        {
            if (value is not null && value.Value <= 0d)
            {
                throw new ArgumentOutOfRangeException(nameof(LineSpacing), value, "Line spacing must be greater than 0.");
            }

            _lineSpacing = value is null ? null : Measure.ValidateLength(value.Value, nameof(LineSpacing));
        }
    }

    /// <summary>
    /// Indentation from the left margin, in points.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not finite.</exception>
    public double? IndentLeft
    {
        get => _indentLeft;
        init => _indentLeft = value is null ? null : Measure.ValidateLength(value.Value, nameof(IndentLeft), allowNegative: true);
    }

    /// <summary>
    /// Indentation from the right margin, in points.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not finite.</exception>
    public double? IndentRight
    {
        get => _indentRight;
        init => _indentRight = value is null ? null : Measure.ValidateLength(value.Value, nameof(IndentRight), allowNegative: true);
    }

    /// <summary>
    /// First-line indentation in points. A positive value indents the first line; a negative value
    /// produces a hanging indent, which WordprocessingML stores as a separate attribute.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not finite.</exception>
    public double? IndentFirstLine
    {
        get => _indentFirstLine;
        init => _indentFirstLine = value is null ? null : Measure.ValidateLength(value.Value, nameof(IndentFirstLine), allowNegative: true);
    }

    /// <summary>
    /// Identifier of the named style to apply, for example a value from <see cref="WordStyleIds"/>.
    /// </summary>
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
    /// Starts the paragraph on a new page. <see langword="false"/> explicitly clears it.
    /// </summary>
    public bool? PageBreakBefore { get; init; }

    /// <summary>
    /// Keeps the paragraph on the same page as the one that follows it, so a heading cannot be
    /// stranded at the bottom of a page. <see langword="false"/> explicitly clears it.
    /// </summary>
    public bool? KeepWithNext { get; init; }

    /// <summary>
    /// Keeps all of the paragraph's lines on one page instead of splitting it across a page break.
    /// <see langword="false"/> explicitly clears it.
    /// </summary>
    public bool? KeepLines { get; init; }

    /// <summary>
    /// Turns the paragraph into a list item at <see cref="ListLevel"/>.
    /// </summary>
    /// <remarks>
    /// The library adds the numbering definition the list needs on first use, the same way it does for
    /// named styles. Use <see cref="Enums.ListStyle.None"/> to remove a paragraph from its list.
    /// </remarks>
    public Enums.ListStyle? ListStyle { get; init; }

    /// <summary>
    /// Nesting depth of a list item, from 0 for the outermost level to 8.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The level is outside 0 to 8.</exception>
    public int? ListLevel
    {
        get => _listLevel;
        init
        {
            if (value is < 0 or > MaxListLevel)
            {
                throw new ArgumentOutOfRangeException(nameof(ListLevel), value, $"List level must be between 0 and {MaxListLevel}.");
            }

            _listLevel = value;
        }
    }

    /// <summary>
    /// <see langword="true"/> when no property is set, so applying it would write nothing.
    /// </summary>
    public bool IsEmpty =>
        Alignment is null
        && SpacingBefore is null
        && SpacingAfter is null
        && LineSpacing is null
        && IndentLeft is null
        && IndentRight is null
        && IndentFirstLine is null
        && StyleId is null
        && PageBreakBefore is null
        && KeepWithNext is null
        && KeepLines is null
        && ListStyle is null
        && ListLevel is null;

    /// <summary>
    /// Layers <paramref name="overrides"/> on top of this format: every property the argument sets
    /// wins, and the ones it leaves unset keep this format's value.
    /// </summary>
    /// <param name="overrides">Format whose set properties take precedence. May be <see langword="null"/>.</param>
    /// <returns>The combined format. Neither input is modified.</returns>
    public ParagraphFormat Merge(ParagraphFormat? overrides)
    {
        if (overrides is null || overrides.IsEmpty)
        {
            return this;
        }

        return new ParagraphFormat
        {
            Alignment = overrides.Alignment ?? Alignment,
            SpacingBefore = overrides.SpacingBefore ?? SpacingBefore,
            SpacingAfter = overrides.SpacingAfter ?? SpacingAfter,
            LineSpacing = overrides.LineSpacing ?? LineSpacing,
            IndentLeft = overrides.IndentLeft ?? IndentLeft,
            IndentRight = overrides.IndentRight ?? IndentRight,
            IndentFirstLine = overrides.IndentFirstLine ?? IndentFirstLine,
            StyleId = overrides.StyleId ?? StyleId,
            PageBreakBefore = overrides.PageBreakBefore ?? PageBreakBefore,
            KeepWithNext = overrides.KeepWithNext ?? KeepWithNext,
            KeepLines = overrides.KeepLines ?? KeepLines,
            ListStyle = overrides.ListStyle ?? ListStyle,
            ListLevel = overrides.ListLevel ?? ListLevel,
        };
    }
}
