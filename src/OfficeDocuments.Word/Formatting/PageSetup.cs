using OfficeDocuments.Word.Enums;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Page size, orientation, and margins for the document.
/// </summary>
/// <remarks>
/// <para>
/// As with the other format records, <see langword="null"/> means "leave this alone". All measurements
/// are in points.
/// </para>
/// <para>
/// This describes a single section. WordprocessingML allows a document to be divided into sections
/// with different page setups; that is deliberately out of scope for now, so a document has one page
/// setup that applies throughout.
/// </para>
/// </remarks>
public sealed record PageSetup
{
    private readonly double? _pageWidth;
    private readonly double? _pageHeight;
    private readonly double? _marginTop;
    private readonly double? _marginBottom;
    private readonly double? _marginLeft;
    private readonly double? _marginRight;
    private readonly double? _headerDistance;
    private readonly double? _footerDistance;

    /// <summary>
    /// A standard paper size, which fills in <see cref="PageWidth"/> and <see cref="PageHeight"/>.
    /// </summary>
    /// <remarks>
    /// Takes precedence over an explicitly set width and height. Set those instead for a custom size.
    /// </remarks>
    public PaperSize? PaperSize { get; init; }

    /// <summary>
    /// Page width in points, for a size <see cref="Enums.PaperSize"/> does not cover.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not positive and finite.</exception>
    public double? PageWidth
    {
        get => _pageWidth;
        init => _pageWidth = value is null ? null : ValidatePositive(value.Value, nameof(PageWidth));
    }

    /// <summary>
    /// Page height in points, for a size <see cref="Enums.PaperSize"/> does not cover.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not positive and finite.</exception>
    public double? PageHeight
    {
        get => _pageHeight;
        init => _pageHeight = value is null ? null : ValidatePositive(value.Value, nameof(PageHeight));
    }

    /// <summary>
    /// Orientation of the page. Landscape swaps the width and height of the chosen size.
    /// </summary>
    public PageOrientation? Orientation { get; init; }

    /// <summary>Top margin in points.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not finite.</exception>
    public double? MarginTop
    {
        get => _marginTop;
        init => _marginTop = value is null ? null : Measure.ValidateLength(value.Value, nameof(MarginTop), allowNegative: true);
    }

    /// <summary>Bottom margin in points.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not finite.</exception>
    public double? MarginBottom
    {
        get => _marginBottom;
        init => _marginBottom = value is null ? null : Measure.ValidateLength(value.Value, nameof(MarginBottom), allowNegative: true);
    }

    /// <summary>Left margin in points.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative or not finite.</exception>
    public double? MarginLeft
    {
        get => _marginLeft;
        init => _marginLeft = value is null ? null : Measure.ValidateLength(value.Value, nameof(MarginLeft));
    }

    /// <summary>Right margin in points.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative or not finite.</exception>
    public double? MarginRight
    {
        get => _marginRight;
        init => _marginRight = value is null ? null : Measure.ValidateLength(value.Value, nameof(MarginRight));
    }

    /// <summary>
    /// Distance from the top of the page to the header, in points.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative or not finite.</exception>
    public double? HeaderDistance
    {
        get => _headerDistance;
        init => _headerDistance = value is null ? null : Measure.ValidateLength(value.Value, nameof(HeaderDistance));
    }

    /// <summary>
    /// Distance from the bottom of the page to the footer, in points.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative or not finite.</exception>
    public double? FooterDistance
    {
        get => _footerDistance;
        init => _footerDistance = value is null ? null : Measure.ValidateLength(value.Value, nameof(FooterDistance));
    }

    /// <summary>
    /// <see langword="true"/> when no property is set, so applying it would write nothing.
    /// </summary>
    public bool IsEmpty =>
        PaperSize is null
        && PageWidth is null
        && PageHeight is null
        && Orientation is null
        && MarginTop is null
        && MarginBottom is null
        && MarginLeft is null
        && MarginRight is null
        && HeaderDistance is null
        && FooterDistance is null;

    /// <summary>
    /// Sets every margin to the same value.
    /// </summary>
    /// <param name="marginInPoints">Margin in points.</param>
    /// <returns>A copy with all four margins set.</returns>
    public PageSetup WithUniformMargins(double marginInPoints)
    {
        return this with
        {
            MarginTop = marginInPoints,
            MarginBottom = marginInPoints,
            MarginLeft = marginInPoints,
            MarginRight = marginInPoints,
        };
    }

    /// <summary>
    /// Layers <paramref name="overrides"/> on top of this setup: every property the argument sets
    /// wins, and the ones it leaves unset keep this setup's value.
    /// </summary>
    /// <param name="overrides">Setup whose set properties take precedence. May be <see langword="null"/>.</param>
    /// <returns>The combined setup. Neither input is modified.</returns>
    public PageSetup Merge(PageSetup? overrides)
    {
        if (overrides is null || overrides.IsEmpty)
        {
            return this;
        }

        return new PageSetup
        {
            PaperSize = overrides.PaperSize ?? PaperSize,
            PageWidth = overrides.PageWidth ?? PageWidth,
            PageHeight = overrides.PageHeight ?? PageHeight,
            Orientation = overrides.Orientation ?? Orientation,
            MarginTop = overrides.MarginTop ?? MarginTop,
            MarginBottom = overrides.MarginBottom ?? MarginBottom,
            MarginLeft = overrides.MarginLeft ?? MarginLeft,
            MarginRight = overrides.MarginRight ?? MarginRight,
            HeaderDistance = overrides.HeaderDistance ?? HeaderDistance,
            FooterDistance = overrides.FooterDistance ?? FooterDistance,
        };
    }

    private static double ValidatePositive(double value, string parameterName)
    {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0d)
        {
            throw new ArgumentOutOfRangeException(parameterName, value, "A page dimension must be a positive finite number of points.");
        }

        return value;
    }
}
