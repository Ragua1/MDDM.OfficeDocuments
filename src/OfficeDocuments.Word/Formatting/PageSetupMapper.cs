using System.Globalization;
using OfficeDocuments.Word.Enums;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Translates between <see cref="PageSetup"/> and the <c>w:sectPr</c> element behind a section.
/// </summary>
internal static class PageSetupMapper
{
    /// <summary>
    /// Standard paper sizes in points, portrait.
    /// </summary>
    private static readonly Dictionary<PaperSize, (double Width, double Height)> PaperSizes = new()
    {
        [PaperSize.A4] = (595.28d, 841.89d),
        [PaperSize.A3] = (841.89d, 1190.55d),
        [PaperSize.A5] = (419.53d, 595.28d),
        [PaperSize.Letter] = (612d, 792d),
        [PaperSize.Legal] = (612d, 1008d),
    };

    /// <summary>
    /// Writes the properties <paramref name="setup"/> sets onto <paramref name="sectionProperties"/>.
    /// </summary>
    internal static void Apply(WordLib.SectionProperties sectionProperties, PageSetup? setup)
    {
        if (setup is null || setup.IsEmpty)
        {
            return;
        }

        ApplyPageSize(sectionProperties, setup);
        ApplyMargins(sectionProperties, setup);
    }

    /// <summary>
    /// Reads back the page setup this library models.
    /// </summary>
    internal static PageSetup Read(WordLib.SectionProperties sectionProperties)
    {
        var size = sectionProperties.GetFirstChild<WordLib.PageSize>();
        var margins = sectionProperties.GetFirstChild<WordLib.PageMargin>();

        var width = TwipsToPoints(size?.Width?.Value);
        var height = TwipsToPoints(size?.Height?.Value);

        return new PageSetup
        {
            PageWidth = width,
            PageHeight = height,
            PaperSize = MatchPaperSize(width, height),
            Orientation = ReadOrientation(size, width, height),
            MarginTop = TwipsToPoints(margins?.Top?.Value),
            MarginBottom = TwipsToPoints(margins?.Bottom?.Value),
            MarginLeft = TwipsToPoints(margins?.Left?.Value),
            MarginRight = TwipsToPoints(margins?.Right?.Value),
            HeaderDistance = TwipsToPoints(margins?.Header?.Value),
            FooterDistance = TwipsToPoints(margins?.Footer?.Value),
        };
    }

    private static void ApplyPageSize(WordLib.SectionProperties sectionProperties, PageSetup setup)
    {
        var dimensions = ResolveDimensions(setup);
        if (dimensions is null && setup.Orientation is null)
        {
            return;
        }

        var size = DataClasses.SectionPropertiesOrderer.GetOrCreate(sectionProperties, () => new WordLib.PageSize());

        if (dimensions is { } resolved)
        {
            var (width, height) = resolved;

            // Landscape is expressed by swapping the stored dimensions as well as setting w:orient;
            // setting only the attribute leaves Word laying the text out on a portrait page.
            if (setup.Orientation == PageOrientation.Landscape && width < height)
            {
                (width, height) = (height, width);
            }
            else if (setup.Orientation == PageOrientation.Portrait && width > height)
            {
                (width, height) = (height, width);
            }

            size.Width = ToTwips(width);
            size.Height = ToTwips(height);
        }

        if (setup.Orientation is { } orientation)
        {
            size.Orient = orientation == PageOrientation.Landscape
                ? WordLib.PageOrientationValues.Landscape
                : WordLib.PageOrientationValues.Portrait;
        }
    }

    private static void ApplyMargins(WordLib.SectionProperties sectionProperties, PageSetup setup)
    {
        if (setup.MarginTop is null
            && setup.MarginBottom is null
            && setup.MarginLeft is null
            && setup.MarginRight is null
            && setup.HeaderDistance is null
            && setup.FooterDistance is null)
        {
            return;
        }

        // Word's defaults, so that setting one margin does not leave the others at zero.
        var margins = DataClasses.SectionPropertiesOrderer.GetOrCreate(sectionProperties, () => new WordLib.PageMargin
        {
            Top = ToSignedTwips(72d),
            Bottom = ToSignedTwips(72d),
            Left = ToTwips(72d),
            Right = ToTwips(72d),
            Header = ToTwips(36d),
            Footer = ToTwips(36d),
            Gutter = 0U,
        });

        // The top and bottom margins are signed in the format, because a negative value is how a
        // header that overlaps the body text is expressed. The others cannot be negative.
        if (setup.MarginTop is { } top)
        {
            margins.Top = ToSignedTwips(top);
        }

        if (setup.MarginBottom is { } bottom)
        {
            margins.Bottom = ToSignedTwips(bottom);
        }

        if (setup.MarginLeft is { } left)
        {
            margins.Left = ToTwips(left);
        }

        if (setup.MarginRight is { } right)
        {
            margins.Right = ToTwips(right);
        }

        if (setup.HeaderDistance is { } header)
        {
            margins.Header = ToTwips(header);
        }

        if (setup.FooterDistance is { } footer)
        {
            margins.Footer = ToTwips(footer);
        }
    }

    private static (double Width, double Height)? ResolveDimensions(PageSetup setup)
    {
        if (setup.PaperSize is { } paperSize && PaperSizes.TryGetValue(paperSize, out var dimensions))
        {
            return dimensions;
        }

        return setup.PageWidth is { } width && setup.PageHeight is { } height ? (width, height) : null;
    }

    private static PaperSize? MatchPaperSize(double? width, double? height)
    {
        if (width is null || height is null)
        {
            return null;
        }

        // Compared with a tolerance because the twip conversion is lossy in both directions.
        var shorterSide = Math.Min(width.Value, height.Value);
        var longerSide = Math.Max(width.Value, height.Value);

        foreach (var (paperSize, dimensions) in PaperSizes)
        {
            if (Math.Abs(shorterSide - dimensions.Width) < 1d && Math.Abs(longerSide - dimensions.Height) < 1d)
            {
                return paperSize;
            }
        }

        return null;
    }

    private static PageOrientation? ReadOrientation(WordLib.PageSize? size, double? width, double? height)
    {
        if (size?.Orient?.Value is { } orient)
        {
            return orient == WordLib.PageOrientationValues.Landscape
                ? PageOrientation.Landscape
                : PageOrientation.Portrait;
        }

        if (width is null || height is null)
        {
            return null;
        }

        return width > height ? PageOrientation.Landscape : PageOrientation.Portrait;
    }

    private static uint ToTwips(double points)
    {
        return (uint)Math.Max(0d, Math.Round(points * 20d, MidpointRounding.AwayFromZero));
    }

    private static int ToSignedTwips(double points)
    {
        return (int)Math.Round(points * 20d, MidpointRounding.AwayFromZero);
    }

    private static double? TwipsToPoints(uint? twips) => twips is null ? null : twips.Value / 20d;

    private static double? TwipsToPoints(int? twips) => twips is null ? null : twips.Value / 20d;

    private static double? TwipsToPoints(string? twips)
    {
        return twips is not null && double.TryParse(twips, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value / 20d
            : null;
    }
}
