using System.ComponentModel;
using System.Xml.Schema;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.DataClasses;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;
using Border = OfficeDocuments.Excel.Styles.Border;
using Font = OfficeDocuments.Excel.Styles.Font;
using Worksheet = OfficeDocuments.Excel.DataClasses.Worksheet;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel;

/// <summary>
/// Class of Spreadsheet
/// </summary>
public partial class Spreadsheet : ISpreadsheet
{
    private readonly List<IWorksheet> _worksheets = [];
    private readonly SpreadsheetDocument _document;
    private readonly bool _isEditable;
    private Style? _defaultStyle;
    private bool _disposed;

    internal WorkbookPart WorkbookPartInternal => _document.WorkbookPart ?? throw new InvalidOperationException();
    private SpreadsheetLib.Workbook WorkbookInternal => WorkbookPartInternal.Workbook ?? throw new InvalidOperationException("The workbook is missing.");
    internal SpreadsheetLib.Sheets SheetsInternal => WorkbookInternal.Sheets ?? throw new InvalidOperationException("The workbook does not contain sheets.");
    internal WorkbookStylesPart WorkbookStylesPartInternal => WorkbookPartInternal.WorkbookStylesPart ?? throw new InvalidOperationException();
    internal SpreadsheetLib.Stylesheet StylesheetInternal => WorkbookStylesPartInternal.Stylesheet ?? InitStylesheet();

    public IReadOnlyList<IWorksheet> Worksheets => _worksheets.AsReadOnly();

    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This property exposes raw OpenXml workbook parts. Prefer using the ISpreadsheet interface methods.")]
    public WorkbookPart WorkbookPart => WorkbookPartInternal;

    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This property exposes raw OpenXml sheets. Prefer using the ISpreadsheet interface methods.")]
    public SpreadsheetLib.Sheets Sheets => SheetsInternal;

    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This property exposes raw OpenXml workbook styles part. Prefer using the ISpreadsheet interface methods.")]
    public WorkbookStylesPart WorkbookStylesPart => WorkbookStylesPartInternal;

    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This property exposes raw OpenXml stylesheet. Prefer using CreateStyle(...) instead.")]
    public SpreadsheetLib.Stylesheet Stylesheet => StylesheetInternal;

    private Spreadsheet(SpreadsheetDocument document, bool createNew, bool isEditable = true)
    {
        _document = document;
        _isEditable = isEditable;

        if (createNew)
        {
            document.AddWorkbookPart();
            WorkbookPartInternal.Workbook = new SpreadsheetLib.Workbook();
            WorkbookPartInternal.Workbook.AppendChild(new SpreadsheetLib.Sheets());
            WorkbookPartInternal.AddNewPart<WorkbookStylesPart>();
            InitStylesheet();
            return;
        }

        if (WorkbookPartInternal.Workbook == null)
        {
            throw new XmlSchemaValidationException("The document is not valid!");
        }

        if (WorkbookPartInternal.WorkbookStylesPart == null)
        {
            WorkbookPartInternal.AddNewPart<WorkbookStylesPart>();
            InitStylesheet();
        }

        foreach (var sheet in SheetsInternal.Elements<SpreadsheetLib.Sheet>())
        {
            if (sheet.Id?.Value is not { Length: > 0 } relationshipId)
            {
                continue;
            }

            var worksheetPart = (WorksheetPart)WorkbookPartInternal.GetPartById(relationshipId);
            var worksheetElement = worksheetPart.Worksheet ?? throw new InvalidOperationException("The worksheet part does not contain a worksheet.");
            var sheetData = worksheetElement.GetFirstChild<SpreadsheetLib.SheetData>() ?? worksheetElement.AppendChild(new SpreadsheetLib.SheetData());
            _worksheets.Add(new Worksheet(this, worksheetPart, sheetData));
        }
    }

    public Spreadsheet(Stream stream, bool createNew = false)
        : this(
            createNew
                ? SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)
                : SpreadsheetDocument.Open(stream, true),
            createNew)
    {
    }

    public Spreadsheet(string filePath, bool createNew = false)
        : this(
            createNew
                ? SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook)
                : SpreadsheetDocument.Open(filePath, true),
            createNew)
    {
    }

    public static ISpreadsheet CreateDocument(Stream stream)
    {
        return new Spreadsheet(SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook), true);
    }

    public static ISpreadsheet OpenDocument(Stream stream, bool isEditable = true)
    {
        return new Spreadsheet(SpreadsheetDocument.Open(stream, isEditable), false, isEditable);
    }

    public IWorksheet AddWorksheet(string? sheetName = null, IStyle? sheetStyle = null)
    {
        var sheetId = SheetsInternal.Elements<SpreadsheetLib.Sheet>().Any()
            ? SheetsInternal.Elements<SpreadsheetLib.Sheet>().Select(sheet => sheet.SheetId?.Value ?? 0U).Max() + 1
            : 1;

        var finalSheetName = sheetName ?? $"Sheet {sheetId}";
        EnsureWorksheetNameAvailable(finalSheetName);

        var sheetData = new SpreadsheetLib.SheetData();
        var worksheetPart = WorkbookPartInternal.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new SpreadsheetLib.Worksheet(sheetData);
        var relationshipId = WorkbookPartInternal.GetIdOfPart(worksheetPart);

        SheetsInternal.Append(new SpreadsheetLib.Sheet
        {
            Id = relationshipId,
            SheetId = sheetId,
            Name = finalSheetName
        });

        var worksheet = new Worksheet(this, worksheetPart, sheetData, _defaultStyle?.CreateMergedStyle(sheetStyle) ?? sheetStyle);
        _worksheets.Add(worksheet);
        return worksheet;
    }

    public IWorksheet? GetWorksheet(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        return _worksheets.FirstOrDefault(worksheet => string.Equals(worksheet.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    public void RenameWorksheet(string currentName, string newName)
    {
        var worksheet = GetWorksheetOrThrow(currentName);
        if (string.IsNullOrWhiteSpace(newName))
        {
            throw new ArgumentException("Worksheet name cannot be null or empty.", nameof(newName));
        }

        if (!string.Equals(currentName, newName, StringComparison.OrdinalIgnoreCase))
        {
            EnsureWorksheetNameAvailable(newName);
        }

        var sheet = GetSheet(worksheet);
        var oldName = sheet.Name?.Value ?? currentName;
        sheet.Name = newName;

        var definedNames = WorkbookPartInternal.Workbook?.DefinedNames;
        if (definedNames != null)
        {
            foreach (var definedName in definedNames.Elements<SpreadsheetLib.DefinedName>())
            {
                if (definedName.Text?.StartsWith($"{oldName}!", StringComparison.Ordinal) == true)
                {
                    definedName.Text = $"{newName}!{definedName.Text[(oldName.Length + 1)..]}";
                }
            }
        }
    }

    public void RemoveWorksheet(string name)
    {
        var worksheet = GetWorksheetOrThrow(name);
        var sheet = GetSheet(worksheet);
        var sheetIndex = GetSheetIndex(sheet);
        AdjustDefinedNamesForRemovedSheet(sheetIndex);

        WorkbookPartInternal.DeletePart(worksheet.WorksheetPart);
        sheet.Remove();
        _worksheets.Remove(worksheet);
    }

    public void MoveWorksheet(string name, uint newPosition)
    {
        var worksheet = GetWorksheetOrThrow(name);
        var sheet = GetSheet(worksheet);
        var sheets = SheetsInternal.Elements<SpreadsheetLib.Sheet>().ToList();
        var oldIndex = sheets.IndexOf(sheet);
        if (oldIndex < 0)
        {
            throw new InvalidOperationException("The worksheet is not part of the workbook.");
        }

        if (newPosition < 1 || newPosition > sheets.Count)
        {
            throw new ArgumentException($"Worksheet position '{newPosition}' is outside the valid range.", nameof(newPosition));
        }

        var newIndex = (int)newPosition - 1;
        if (oldIndex == newIndex)
        {
            return;
        }

        sheet.Remove();

        var remainingSheets = SheetsInternal.Elements<SpreadsheetLib.Sheet>().ToList();
        if (newIndex >= remainingSheets.Count)
        {
            SheetsInternal.Append(sheet);
        }
        else
        {
            SheetsInternal.InsertBefore(sheet, remainingSheets[newIndex]);
        }

        AdjustDefinedNamesForMovedSheet(oldIndex, newIndex);
        RebuildWorksheetOrder();
    }

    public IWorksheet CopyWorksheet(string sourceName, string? newName = null)
    {
        var sourceWorksheet = GetWorksheetOrThrow(sourceName);
        var sourcePart = sourceWorksheet.WorksheetPart;
        var sourceSheet = GetSheet(sourceWorksheet);
        var finalName = newName ?? $"{sourceWorksheet.Name} Copy";
        EnsureWorksheetNameAvailable(finalName);

        var worksheetPart = WorkbookPartInternal.AddNewPart<WorksheetPart>();
        var clonedWorksheet = (SpreadsheetLib.Worksheet)(sourcePart.Worksheet?.CloneNode(true) ?? throw new InvalidOperationException("The worksheet part does not contain a worksheet."));
        clonedWorksheet.RemoveAllChildren<SpreadsheetLib.Hyperlinks>();
        clonedWorksheet.RemoveAllChildren<SpreadsheetLib.LegacyDrawing>();
        clonedWorksheet.RemoveAllChildren<SpreadsheetLib.TableParts>();
        worksheetPart.Worksheet = clonedWorksheet;

        var relationshipId = WorkbookPartInternal.GetIdOfPart(worksheetPart);
        var newSheetId = SheetsInternal.Elements<SpreadsheetLib.Sheet>().Select(sheet => sheet.SheetId?.Value ?? 0U).DefaultIfEmpty().Max() + 1;
        var newSheet = new SpreadsheetLib.Sheet
        {
            Id = relationshipId,
            SheetId = newSheetId,
            Name = finalName
        };

        SheetsInternal.InsertAfter(newSheet, sourceSheet);
        var sheetData = clonedWorksheet.GetFirstChild<SpreadsheetLib.SheetData>() ?? clonedWorksheet.AppendChild(new SpreadsheetLib.SheetData());
        var worksheet = new Worksheet(this, worksheetPart, sheetData);
        RebuildWorksheetOrder(worksheet);
        return worksheet;
    }

    public void SetWorksheetHidden(string name, bool isHidden)
    {
        var sheet = GetSheet(GetWorksheetOrThrow(name));
        sheet.State = isHidden ? SpreadsheetLib.SheetStateValues.Hidden : SpreadsheetLib.SheetStateValues.Visible;
    }

    public IEnumerable<string> GetWorksheetsName()
    {
        return SheetsInternal.Elements<SpreadsheetLib.Sheet>()
            .Select(sheet => sheet.Name?.Value ?? throw new InvalidOperationException("Worksheet name is missing."))
            .ToArray();
    }

    public void Close()
    {
        if (_isEditable)
        {
            WorkbookPartInternal.Workbook?.Save();
        }

        if (!_disposed)
        {
            _document.Dispose();
        }

        _disposed = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
        {
            return;
        }

        if (disposing)
        {
            Close();
        }

        _disposed = true;
    }

    ~Spreadsheet()
    {
        Dispose(false);
    }

    public SpreadsheetLib.Stylesheet InitStylesheet()
    {
        var stylesheet = WorkbookStylesPartInternal.Stylesheet = new SpreadsheetLib.Stylesheet();
        stylesheet.CellFormats = new SpreadsheetLib.CellFormats();
        stylesheet.Fills = new SpreadsheetLib.Fills(
            new SpreadsheetLib.Fill { PatternFill = new SpreadsheetLib.PatternFill { PatternType = SpreadsheetLib.PatternValues.None } },
            new SpreadsheetLib.Fill { PatternFill = new SpreadsheetLib.PatternFill { PatternType = SpreadsheetLib.PatternValues.Gray125 } }
        );

        _defaultStyle = new Style(
            stylesheet,
            font: new Font { FontSize = 11, Color = Color.Black, FontName = FontNameValues.Calibri },
            fill: null,
            border: new Border(),
            numberFormat: null,
            alignment: null);

        stylesheet.CellStyleFormats = new SpreadsheetLib.CellStyleFormats(_defaultStyle.ElementInternal.CloneNode(true));
        return stylesheet;
    }

    internal string GetWorksheetName(Worksheet worksheet) => GetSheet(worksheet).Name?.Value ?? throw new InvalidOperationException("Worksheet name is missing.");

    internal bool IsWorksheetHidden(Worksheet worksheet)
    {
        var state = GetSheet(worksheet).State?.Value;
        return state == SpreadsheetLib.SheetStateValues.Hidden || state == SpreadsheetLib.SheetStateValues.VeryHidden;
    }

    private Worksheet GetWorksheetOrThrow(string name)
    {
        return GetWorksheet(name) as Worksheet ?? throw new ArgumentException($"Cannot find worksheet with name '{name}'.", nameof(name));
    }

    private SpreadsheetLib.Sheet GetSheet(Worksheet worksheet)
    {
        var relationshipId = WorkbookPartInternal.GetIdOfPart(worksheet.WorksheetPart);
        return SheetsInternal.Elements<SpreadsheetLib.Sheet>().First(sheet => sheet.Id == relationshipId);
    }

    private int GetSheetIndex(SpreadsheetLib.Sheet sheet)
    {
        var sheets = SheetsInternal.Elements<SpreadsheetLib.Sheet>().ToList();
        var index = sheets.IndexOf(sheet);
        if (index < 0)
        {
            throw new InvalidOperationException("The worksheet is not part of the workbook.");
        }

        return index;
    }

    private void RebuildWorksheetOrder(params Worksheet[] additionalWorksheets)
    {
        var allWorksheets = _worksheets.OfType<Worksheet>().Concat(additionalWorksheets).Distinct().ToDictionary(
            worksheet => WorkbookPartInternal.GetIdOfPart(worksheet.WorksheetPart),
            worksheet => (IWorksheet)worksheet);

        _worksheets.Clear();
        foreach (var sheet in SheetsInternal.Elements<SpreadsheetLib.Sheet>())
        {
            if (sheet.Id?.Value is { Length: > 0 } relationshipId && allWorksheets.TryGetValue(relationshipId, out var worksheet))
            {
                _worksheets.Add(worksheet);
            }
        }
    }

    private void EnsureWorksheetNameAvailable(string name)
    {
        WorksheetNameValidator.Validate(name, nameof(name));

        if (GetWorksheet(name) != null)
        {
            throw new ArgumentException($"Worksheet '{name}' already exists.", nameof(name));
        }
    }

    private void AdjustDefinedNamesForRemovedSheet(int removedIndex)
    {
        var definedNames = WorkbookPartInternal.Workbook?.DefinedNames;
        if (definedNames == null)
        {
            return;
        }

        foreach (var definedName in definedNames.Elements<SpreadsheetLib.DefinedName>().ToList())
        {
            if (definedName.LocalSheetId?.Value == removedIndex)
            {
                definedName.Remove();
                continue;
            }

            if (definedName.LocalSheetId?.Value > removedIndex)
            {
                definedName.LocalSheetId = definedName.LocalSheetId.Value - 1;
            }
        }
    }

    private void AdjustDefinedNamesForMovedSheet(int oldIndex, int newIndex)
    {
        var definedNames = WorkbookPartInternal.Workbook?.DefinedNames;
        if (definedNames == null)
        {
            return;
        }

        foreach (var definedName in definedNames.Elements<SpreadsheetLib.DefinedName>())
        {
            if (definedName.LocalSheetId == null)
            {
                continue;
            }

            var currentIndex = (int)definedName.LocalSheetId.Value;
            if (currentIndex == oldIndex)
            {
                definedName.LocalSheetId = Convert.ToUInt32(newIndex);
            }
            else if (oldIndex < newIndex && currentIndex > oldIndex && currentIndex <= newIndex)
            {
                definedName.LocalSheetId = Convert.ToUInt32(currentIndex - 1);
            }
            else if (oldIndex > newIndex && currentIndex >= newIndex && currentIndex < oldIndex)
            {
                definedName.LocalSheetId = Convert.ToUInt32(currentIndex + 1);
            }
        }
    }

}
