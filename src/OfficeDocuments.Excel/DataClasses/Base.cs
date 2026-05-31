using OfficeDocuments.Excel.Interfaces;

namespace OfficeDocuments.Excel.DataClasses;

internal abstract class Base : IBase
{
    public IWorksheet Worksheet { get; protected set; } = default!;
    public IStyle? Style { get; protected set; }
    protected Worksheet OwnerWorksheet => (Worksheet)Worksheet;
    protected Spreadsheet OwnerSpreadsheet => OwnerWorksheet.Spreadsheet;

    // Ctor used when the Worksheet isn't available yet (e.g., constructing Worksheet itself)
    protected Base(IStyle? cellStyle)
    {
        // Worksheet will be set by derived type after construction
        MergeStyles(cellStyle);
    }

    protected Base(IWorksheet? worksheet, IStyle? cellStyle = null)
    {
        Worksheet = worksheet!; // set by derived types when passing null (e.g., Worksheet itself)
        MergeStyles(cellStyle);
    }
    protected Base(IWorksheet worksheet, uint cellStyle)
    {
        Worksheet = worksheet;

        if (cellStyle > 0)
        {
            Style = new Style(OwnerSpreadsheet.StylesheetInternal, cellStyle);
        }
    }

    protected IStyle? MergeStyles(params IStyle?[] styles)
    {
        foreach (var style in styles.Where(s => s != null))
        {
            Style = Style?.CreateMergedStyle(style) ?? style;
        }

        return Style;
    }

    public virtual IStyle? AddStyle(params IStyle?[] styles)
    {
        return MergeStyles(styles);
    }
}