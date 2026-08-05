using OfficeDocuments.Excel.Options;

namespace OfficeDocuments.Excel.DataClasses;

internal partial class Worksheet
{
    private DataValidationWriter? _dataValidationWriter;
    private DataValidationWriter DataValidationWriter => _dataValidationWriter ??= new DataValidationWriter(WorksheetElement, ElementOrderer);

    private ConditionalFormattingWriter? _conditionalFormattingWriter;
    private ConditionalFormattingWriter ConditionalFormattingWriter =>
        _conditionalFormattingWriter ??= new ConditionalFormattingWriter(WorksheetElement, ElementOrderer, Spreadsheet.GetOrCreateDifferentialFormat);

    private HyperlinkStore? _hyperlinkStore;
    private HyperlinkStore HyperlinkStore => _hyperlinkStore ??= new HyperlinkStore(WorksheetPart, WorksheetElement, ElementOrderer);

    internal void AddDataValidation(string reference, DataValidationOptions options) => DataValidationWriter.Add(reference, options);

    internal void AddConditionalFormatting(string reference, ConditionalFormattingOptions options) => ConditionalFormattingWriter.Add(reference, options);

    internal void SetCellHyperlink(Cell cell, string target, string? displayText)
    {
        if (string.IsNullOrWhiteSpace(target))
        {
            throw new ArgumentException("Hyperlink target cannot be null or empty.", nameof(target));
        }

        if (!string.IsNullOrEmpty(displayText))
        {
            cell.SetValue(displayText);
        }

        HyperlinkStore.Set(cell.CellReference, target);
    }

    internal string? GetCellHyperlink(string cellReference) => HyperlinkStore.Get(cellReference);
}
