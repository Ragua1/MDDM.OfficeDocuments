using OfficeDocuments.Word.Enums;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Builds the numbering definitions that turn a paragraph into a list item.
/// </summary>
/// <remarks>
/// <para>
/// A list in WordprocessingML is not a property of the paragraph. The paragraph only carries a
/// <c>w:numPr</c> pointing at a <c>w:num</c>, which points at an <c>w:abstractNum</c> in a separate
/// numbering part, and that is where the bullet character, the number format, and the indentation of
/// every level actually live. Without those definitions a list item renders as an ordinary paragraph.
/// </para>
/// <para>
/// Each definition declares all nine levels up front, because a level referenced but not defined
/// falls back to the document default rather than to something list-shaped.
/// </para>
/// </remarks>
internal static class ListNumbering
{
    /// <summary>
    /// Numbering identifier that means "not in a list". Reserved by the format, not by this library.
    /// </summary>
    internal const int NoNumberingId = 0;

    /// <summary>
    /// Indentation step per nesting level, in twips. 720 is half an inch, which is Word's own step.
    /// </summary>
    private const int IndentStepTwips = 720;

    /// <summary>
    /// Distance the text hangs back from the marker, in twips.
    /// </summary>
    private const int HangingIndentTwips = 360;

    /// <summary>
    /// Bullet glyphs by depth, cycling. Plain Unicode rather than the private-use characters Word
    /// pairs with the Symbol font, so the marker survives in readers that lack that font.
    /// </summary>
    private static readonly string[] BulletGlyphs = ["•", "◦", "▪"];

    /// <summary>
    /// Builds an abstract numbering definition for <paramref name="style"/>.
    /// </summary>
    internal static WordLib.AbstractNum CreateAbstractNumbering(ListStyle style, int abstractNumberingId)
    {
        var abstractNumbering = new WordLib.AbstractNum
        {
            AbstractNumberId = abstractNumberingId,
            MultiLevelType = new WordLib.MultiLevelType { Val = WordLib.MultiLevelValues.HybridMultilevel },
        };

        for (var level = 0; level <= ParagraphFormat.MaxListLevel; level++)
        {
            abstractNumbering.AppendChild(CreateLevel(style, level));
        }

        return abstractNumbering;
    }

    /// <summary>
    /// Classifies an abstract numbering definition by the format of its first level, so a list read
    /// from a document the library did not write can still be reported as a bullet or number list.
    /// </summary>
    internal static ListStyle? ClassifyAbstractNumbering(WordLib.AbstractNum abstractNumbering)
    {
        var firstLevel = abstractNumbering.Elements<WordLib.Level>()
            .FirstOrDefault(level => level.LevelIndex?.Value is null or 0);

        var format = firstLevel?.NumberingFormat?.Val?.Value;
        if (format is null)
        {
            return null;
        }

        if (format == WordLib.NumberFormatValues.Bullet)
        {
            return ListStyle.Bullet;
        }

        if (format == WordLib.NumberFormatValues.Decimal
            || format == WordLib.NumberFormatValues.LowerLetter
            || format == WordLib.NumberFormatValues.UpperLetter
            || format == WordLib.NumberFormatValues.LowerRoman
            || format == WordLib.NumberFormatValues.UpperRoman)
        {
            return ListStyle.Number;
        }

        return null;
    }

    private static WordLib.Level CreateLevel(ListStyle style, int level)
    {
        // Typed setters again, because CT_Lvl has its own strict child sequence.
        var definition = new WordLib.Level
        {
            LevelIndex = level,
            StartNumberingValue = new WordLib.StartNumberingValue { Val = 1 },
            NumberingFormat = new WordLib.NumberingFormat { Val = GetNumberFormat(style, level) },
            LevelText = new WordLib.LevelText { Val = GetLevelText(style, level) },
            LevelJustification = new WordLib.LevelJustification { Val = WordLib.LevelJustificationValues.Left },
            PreviousParagraphProperties = new WordLib.PreviousParagraphProperties
            {
                Indentation = new WordLib.Indentation
                {
                    Left = (IndentStepTwips * (level + 1)).ToString(System.Globalization.CultureInfo.InvariantCulture),
                    Hanging = HangingIndentTwips.ToString(System.Globalization.CultureInfo.InvariantCulture),
                },
            },
        };

        return definition;
    }

    private static WordLib.NumberFormatValues GetNumberFormat(ListStyle style, int level)
    {
        if (style == ListStyle.Bullet)
        {
            return WordLib.NumberFormatValues.Bullet;
        }

        // Word's own multilevel default alternates the format by depth, which keeps nested lists
        // readable: 1. then a. then i.
        return (level % 3) switch
        {
            0 => WordLib.NumberFormatValues.Decimal,
            1 => WordLib.NumberFormatValues.LowerLetter,
            _ => WordLib.NumberFormatValues.LowerRoman,
        };
    }

    private static string GetLevelText(ListStyle style, int level)
    {
        // "%n" is a placeholder for the counter of level n, numbered from 1.
        return style == ListStyle.Bullet
            ? BulletGlyphs[level % BulletGlyphs.Length]
            : $"%{level + 1}.";
    }
}
