using DocumentFormat.OpenXml;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// Places children of <c>w:sectPr</c> at their schema position.
/// </summary>
/// <remarks>
/// <para>
/// <c>CT_SectPr</c> is a fixed sequence, and unlike most elements this library writes it has no typed
/// SDK properties to lean on — the SDK models these children as a plain collection. So the order has
/// to be applied deliberately: header and footer references come first, then the page size, then the
/// margins, with the title-page switch much later.
/// </para>
/// <para>
/// Getting this wrong produces a document that round-trips through this library perfectly and then
/// makes Word offer to repair it, which is the same defect class that reached disk in the Excel module
/// before its schema-validation gate existed.
/// </para>
/// </remarks>
internal static class SectionPropertiesOrderer
{
    /// <summary>
    /// Inserts <paramref name="element"/> into <paramref name="sectionProperties"/> in schema order.
    /// </summary>
    internal static void Insert(WordLib.SectionProperties sectionProperties, OpenXmlElement element)
    {
        var rank = RankOf(element);

        foreach (var existing in sectionProperties.ChildElements)
        {
            if (RankOf(existing) > rank)
            {
                sectionProperties.InsertBefore(element, existing);
                return;
            }
        }

        sectionProperties.AppendChild(element);
    }

    /// <summary>
    /// Returns the existing child of type <typeparamref name="T"/>, or adds one in schema order.
    /// </summary>
    internal static T GetOrCreate<T>(WordLib.SectionProperties sectionProperties, Func<T> create)
        where T : OpenXmlElement
    {
        var existing = sectionProperties.GetFirstChild<T>();
        if (existing is not null)
        {
            return existing;
        }

        var element = create();
        Insert(sectionProperties, element);

        return element;
    }

    /// <summary>
    /// Position of a child in the <c>CT_SectPr</c> sequence.
    /// </summary>
    /// <remarks>
    /// Anything not listed sorts last and keeps its relative position, so an element this library does
    /// not write is never moved.
    /// </remarks>
    private static int RankOf(OpenXmlElement element)
    {
        return element switch
        {
            WordLib.HeaderReference or WordLib.FooterReference => 0,
            WordLib.FootnoteProperties => 1,
            WordLib.EndnoteProperties => 2,
            WordLib.SectionType => 3,
            WordLib.PageSize => 4,
            WordLib.PageMargin => 5,
            WordLib.PaperSource => 6,
            WordLib.PageBorders => 7,
            WordLib.LineNumberType => 8,
            WordLib.PageNumberType => 9,
            WordLib.Columns => 10,
            WordLib.FormProtection => 11,
            WordLib.VerticalTextAlignmentOnPage => 12,
            WordLib.NoEndnote => 13,
            WordLib.TitlePage => 14,
            WordLib.TextDirection => 15,
            WordLib.BiDi => 16,
            WordLib.GutterOnRight => 17,
            WordLib.DocGrid => 18,
            WordLib.PrinterSettingsReference => 19,
            _ => 100,
        };
    }
}
