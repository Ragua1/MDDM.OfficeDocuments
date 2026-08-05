using OfficeDocuments.Word.Enums;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Character-level formatting of a text run: bold, italic, underline, font, size, and colour.
/// </summary>
/// <remarks>
/// <para>
/// Every property is optional, and <see langword="null"/> means "leave this alone" rather than
/// "turn it off". That distinction matters in WordprocessingML: <c>&lt;w:b/&gt;</c> switches bold on,
/// <c>&lt;w:b w:val="0"/&gt;</c> switches it off, and writing neither lets the paragraph's style
/// decide. Setting <see cref="Bold"/> to <see langword="false"/> therefore actively overrides a bold
/// style, which is not the same as not setting it.
/// </para>
/// <para>
/// Being a record, a format is cheap to vary and reuse:
/// </para>
/// <code>
/// var body = new TextFormat { FontName = "Calibri", FontSize = 11 };
/// paragraph.AddText("normal", body);
/// paragraph.AddText("emphasis", body with { Bold = true });
/// </code>
/// </remarks>
public sealed record TextFormat
{
    private readonly double? _fontSize;
    private readonly string? _fontName;
    private readonly string? _color;
    private readonly string? _styleId;

    /// <summary>
    /// Identifier of a character style to apply, for example <see cref="WordStyleIds.Hyperlink"/>.
    /// </summary>
    /// <remarks>
    /// A character style is the reusable counterpart of the direct formatting on this record: the
    /// other properties override whatever the style sets.
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
    /// Bold. <see langword="false"/> explicitly clears bold inherited from a style.
    /// </summary>
    public bool? Bold { get; init; }

    /// <summary>
    /// Italic. <see langword="false"/> explicitly clears italic inherited from a style.
    /// </summary>
    public bool? Italic { get; init; }

    /// <summary>
    /// Underline style. Use <see cref="UnderlineType.None"/> to clear an inherited underline.
    /// </summary>
    public UnderlineType? Underline { get; init; }

    /// <summary>
    /// Strikethrough. <see langword="false"/> explicitly clears it.
    /// </summary>
    public bool? Strikethrough { get; init; }

    /// <summary>
    /// Renders the text in capitals without changing the stored characters.
    /// <see langword="false"/> explicitly clears it.
    /// </summary>
    public bool? AllCaps { get; init; }

    /// <summary>
    /// Renders lower-case letters as smaller capitals. <see langword="false"/> explicitly clears it.
    /// </summary>
    public bool? SmallCaps { get; init; }

    /// <summary>
    /// Highlight colour. Use <see cref="HighlightColor.None"/> to clear an inherited highlight.
    /// </summary>
    /// <remarks>
    /// Distinct from <see cref="Color"/>: this is the marker-pen background, and WordprocessingML
    /// restricts it to a fixed palette rather than accepting an arbitrary value.
    /// </remarks>
    public HighlightColor? Highlight { get; init; }

    /// <summary>
    /// Superscript or subscript. Use <see cref="TextVerticalPosition.Baseline"/> to clear it.
    /// </summary>
    public TextVerticalPosition? VerticalPosition { get; init; }

    /// <summary>
    /// Font family name, for example <c>Calibri</c>.
    /// </summary>
    /// <exception cref="ArgumentException">The name is empty or whitespace.</exception>
    public string? FontName
    {
        get => _fontName;
        init
        {
            if (value is not null)
            {
                ArgumentException.ThrowIfNullOrWhiteSpace(value, nameof(FontName));
            }

            _fontName = value;
        }
    }

    /// <summary>
    /// Font size in points. Stored as half-points, so a half-point is the finest usable step.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The size is outside what Word can store.</exception>
    public double? FontSize
    {
        get => _fontSize;
        init => _fontSize = value is null ? null : Measure.ValidateFontSize(value.Value, nameof(FontSize));
    }

    /// <summary>
    /// Text colour as 6 hex digits (<c>RRGGBB</c>), optionally <c>#</c>-prefixed, or <c>auto</c>.
    /// Normalized on assignment, so an invalid colour fails here rather than in Word.
    /// </summary>
    /// <exception cref="ArgumentException">The value is not a colour Word can store.</exception>
    public string? Color
    {
        get => _color;
        init => _color = value is null ? null : HexColor.Normalize(value, nameof(Color));
    }

    /// <summary>
    /// <see langword="true"/> when no property is set, so applying it would write nothing.
    /// </summary>
    public bool IsEmpty =>
        StyleId is null
        && Bold is null
        && Italic is null
        && Underline is null
        && Strikethrough is null
        && AllCaps is null
        && SmallCaps is null
        && Highlight is null
        && VerticalPosition is null
        && FontName is null
        && FontSize is null
        && Color is null;

    /// <summary>
    /// Layers <paramref name="overrides"/> on top of this format: every property the argument sets
    /// wins, and the ones it leaves unset keep this format's value.
    /// </summary>
    /// <param name="overrides">Format whose set properties take precedence. May be <see langword="null"/>.</param>
    /// <returns>The combined format. Neither input is modified.</returns>
    public TextFormat Merge(TextFormat? overrides)
    {
        if (overrides is null || overrides.IsEmpty)
        {
            return this;
        }

        return new TextFormat
        {
            StyleId = overrides.StyleId ?? StyleId,
            Bold = overrides.Bold ?? Bold,
            Italic = overrides.Italic ?? Italic,
            Underline = overrides.Underline ?? Underline,
            Strikethrough = overrides.Strikethrough ?? Strikethrough,
            AllCaps = overrides.AllCaps ?? AllCaps,
            SmallCaps = overrides.SmallCaps ?? SmallCaps,
            Highlight = overrides.Highlight ?? Highlight,
            VerticalPosition = overrides.VerticalPosition ?? VerticalPosition,
            FontName = overrides.FontName ?? FontName,
            FontSize = overrides.FontSize ?? FontSize,
            Color = overrides.Color ?? Color,
        };
    }
}
