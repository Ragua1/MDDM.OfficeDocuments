using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Xml.Schema;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.DataClasses;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Extensions;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Options;
using Alignment = OfficeDocuments.Excel.Styles.Alignment;
using Border = OfficeDocuments.Excel.Styles.Border;
using Fill = OfficeDocuments.Excel.Styles.Fill;
using Font = OfficeDocuments.Excel.Styles.Font;
using NumberingFormat = OfficeDocuments.Excel.Styles.NumberingFormat;
using Worksheet = OfficeDocuments.Excel.DataClasses.Worksheet;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel;

/// <summary>
/// Class of Spreadsheet
/// </summary>
public class Spreadsheet : ISpreadsheet
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

        var worksheet = new Worksheet(this, worksheetPart, sheetData, _defaultStyle?.CreateMergedStyle(sheetStyle));
        _worksheets.Add(worksheet);
        return worksheet;
    }

    public IStyle CreateStyle(Font? font = null, Fill? fill = null, Border? border = null, NumberingFormat? numberFormat = null, Alignment? alignment = null)
    {
        return new Style(StylesheetInternal, font, fill, border, numberFormat, alignment);
    }

    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This overload exposes raw OpenXml stylesheet plumbing. Prefer CreateStyle(...) without a Stylesheet parameter.")]
    public IStyle CreateStyle(SpreadsheetLib.Stylesheet stylesheet, Font? font = null, Fill? fill = null, Border? border = null, NumberingFormat? numberFormat = null, Alignment? alignment = null)
    {
        ArgumentNullException.ThrowIfNull(stylesheet);
        return new Style(stylesheet, font, fill, border, numberFormat, alignment);
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

    public ITableInfo AddTable(string worksheetName, ICell startCell, ICell endCell, List<string> columnsName)
    {
        return AddTableCore(worksheetName, startCell, endCell, columnsName, options: null);
    }

    public ITableInfo AddTable(IRange range, List<string> columnsName, TableCreateOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(range);

        var startCell = range.Worksheet.AddCellOnIndex(range.FromColumn, range.FromRow);
        var endCell = range.Worksheet.AddCellOnIndex(range.ToColumn, range.ToRow);
        return AddTableCore(range.Worksheet.Name, startCell, endCell, columnsName, options);
    }

    public ITableInfo? GetTable(string worksheetName, string tableName)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);
        ArgumentException.ThrowIfNullOrEmpty(tableName);

        var worksheet = GetWorksheetOrThrow(worksheetName);
        var tablePart = FindTableDefinitionPart(worksheet.WorksheetPart, tableName);
        return tablePart == null ? null : ToTableInfo(worksheetName, tablePart);
    }

    public IEnumerable<ITableInfo> GetTables(string worksheetName)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);

        var worksheet = GetWorksheetOrThrow(worksheetName);
        return worksheet.WorksheetPart.TableDefinitionParts
            .Select(part => ToTableInfo(worksheetName, part))
            .ToArray();
    }

    public IEnumerable<ITableInfo> GetTables()
    {
        return _worksheets.OfType<Worksheet>()
            .SelectMany(ws => ws.WorksheetPart.TableDefinitionParts
                .Select(part => ToTableInfo(ws.Name, part)))
            .ToArray();
    }

    public void RenameTable(string worksheetName, string tableName, string newName)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);
        ArgumentException.ThrowIfNullOrEmpty(tableName);
        ArgumentException.ThrowIfNullOrEmpty(newName);

        var allTables = GetTables();
        if (allTables.Any(t => string.Equals(t.Name, newName, StringComparison.OrdinalIgnoreCase) && !string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase)))
        {
            throw new ArgumentException($"A table named '{newName}' already exists in the workbook.", nameof(newName));
        }

        var worksheet = GetWorksheetOrThrow(worksheetName);
        var tablePart = FindTableDefinitionPart(worksheet.WorksheetPart, tableName)
            ?? throw new ArgumentException($"Table '{tableName}' not found on worksheet '{worksheetName}'.", nameof(tableName));
        var table = tablePart.Table ?? throw new InvalidOperationException($"Table '{tableName}' does not contain a table definition.");

        table.Name = newName;
        table.DisplayName = newName;
    }

    public void ResizeTable(string worksheetName, string tableName, IRange newRange)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);
        ArgumentException.ThrowIfNullOrEmpty(tableName);
        ArgumentNullException.ThrowIfNull(newRange);

        var worksheet = GetWorksheetOrThrow(worksheetName);
        var tablePart = FindTableDefinitionPart(worksheet.WorksheetPart, tableName)
            ?? throw new ArgumentException($"Table '{tableName}' not found on worksheet '{worksheetName}'.", nameof(tableName));

        var table = tablePart.Table ?? throw new InvalidOperationException($"Table '{tableName}' does not contain a table definition.");
        var existingColumnCount = (int)(table.TableColumns?.Count?.Value ?? 0);
        var newColumnCount = (int)(newRange.ToColumn - newRange.FromColumn + 1);
        if (newColumnCount != existingColumnCount)
        {
            throw new ArgumentException($"Cannot resize table '{tableName}': new range has {newColumnCount} columns but table has {existingColumnCount} columns.", nameof(newRange));
        }

        var startRef = CellExtension.GetExcelCellReference(newRange.FromColumn, newRange.FromRow);
        var endRef = CellExtension.GetExcelCellReference(newRange.ToColumn, newRange.ToRow);
        var newRef = $"{startRef}:{endRef}";
        table.Reference = newRef;
        if (table.AutoFilter != null)
        {
            table.AutoFilter.Reference = newRef;
        }
    }

    public void RemoveTable(string worksheetName, string tableName)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);
        ArgumentException.ThrowIfNullOrEmpty(tableName);

        var worksheet = GetWorksheetOrThrow(worksheetName);
        var tablePart = FindTableDefinitionPart(worksheet.WorksheetPart, tableName)
            ?? throw new ArgumentException($"Table '{tableName}' not found on worksheet '{worksheetName}'.", nameof(tableName));

        var tableRelId = worksheet.WorksheetPart.GetIdOfPart(tablePart);
        var tableParts = worksheet.WorksheetElement.GetFirstChild<SpreadsheetLib.TableParts>();
        if (tableParts != null)
        {
            var tablePartRef = tableParts.Elements<SpreadsheetLib.TablePart>()
                .FirstOrDefault(tp => tp.Id?.Value == tableRelId);
            tablePartRef?.Remove();

            if (!tableParts.Elements<SpreadsheetLib.TablePart>().Any())
            {
                tableParts.Remove();
            }
            else
            {
                tableParts.Count = Convert.ToUInt32(tableParts.Elements<SpreadsheetLib.TablePart>().Count());
            }
        }

        worksheet.WorksheetPart.DeletePart(tablePart);
    }

    public void AddNamedRange(string name, IRange range, bool worksheetScoped = false)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        ArgumentNullException.ThrowIfNull(range);

        if (!IsValidNamedRange(name))
        {
            throw new ArgumentException($"Named range '{name}' is not valid.", nameof(name));
        }

        var definedNames = WorkbookPartInternal.Workbook?.DefinedNames ?? WorkbookPartInternal.Workbook!.AppendChild(new SpreadsheetLib.DefinedNames());
        var localSheetId = worksheetScoped ? Convert.ToUInt32(GetSheetIndex(GetSheet(GetWorksheetOrThrow(range.Worksheet.Name)))) : (uint?)null;

        if (definedNames.Elements<SpreadsheetLib.DefinedName>().Any(definedName =>
                string.Equals(definedName.Name?.Value, name, StringComparison.OrdinalIgnoreCase)
                && (definedName.LocalSheetId == null && localSheetId == null || definedName.LocalSheetId?.Value == localSheetId)))
        {
            throw new ArgumentException($"Named range '{name}' already exists.", nameof(name));
        }

        definedNames.Append(new SpreadsheetLib.DefinedName
        {
            Name = name,
            LocalSheetId = localSheetId,
            Text = $"{range.Worksheet.Name}!{range.Reference}"
        });
    }

    public void ProtectWorkbook(string? password = null)
    {
        var workbookProtection = WorkbookPartInternal.Workbook?.GetFirstChild<SpreadsheetLib.WorkbookProtection>();
        if (workbookProtection == null)
        {
            workbookProtection = WorkbookPartInternal.Workbook!.AppendChild(new SpreadsheetLib.WorkbookProtection());
        }

        workbookProtection.LockStructure = true;
        if (!string.IsNullOrEmpty(password))
        {
            workbookProtection.WorkbookPassword = ComputeProtectionPassword(password);
        }
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

    internal uint GetOrCreateDifferentialFormat(IStyle style)
    {
        var differentialFormats = StylesheetInternal.DifferentialFormats ??= new SpreadsheetLib.DifferentialFormats();
        var differentialFormat = CreateDifferentialFormat(style);
        var existingFormats = differentialFormats.Elements<SpreadsheetLib.DifferentialFormat>().ToList();
        var existingIndex = existingFormats.FindIndex(existing => Utils.OpenXmlElementsEqual(existing, differentialFormat));
        if (existingIndex >= 0)
        {
            return Convert.ToUInt32(existingIndex);
        }

        differentialFormats.Append(differentialFormat);
        differentialFormats.Count = Convert.ToUInt32(differentialFormats.Count());
        return Convert.ToUInt32(existingFormats.Count);
    }

    internal static HexBinaryValue ComputeProtectionPassword(string password)
    {
        var hash = 0;
        for (var index = password.Length - 1; index >= 0; index--)
        {
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
            hash ^= password[index];
        }

        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
        hash ^= password.Length;
        hash ^= 0xCE4B;

        return new HexBinaryValue(hash.ToString("X4", CultureInfo.InvariantCulture));
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
        if (GetWorksheet(name) != null)
        {
            throw new ArgumentException($"Worksheet '{name}' already exists.", nameof(name));
        }
    }

    private static bool IsValidNamedRange(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        if (!char.IsLetter(name[0]) && name[0] != '_')
        {
            return false;
        }

        if (name.Any(character => !(char.IsLetterOrDigit(character) || character is '_' or '.')))
        {
            return false;
        }

        return !name.TryGetExcelRange(out _);
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

    private SpreadsheetLib.DifferentialFormat CreateDifferentialFormat(IStyle style)
    {
        var differentialFormat = new SpreadsheetLib.DifferentialFormat();

        if (style.FontId > 0)
        {
            var font = StylesheetInternal.Fonts?.Elements<SpreadsheetLib.Font>().ElementAt(style.FontId);
            if (font != null)
            {
                differentialFormat.Font = (SpreadsheetLib.Font)font.CloneNode(true);
            }
        }

        if (style.FillId > 0)
        {
            var fill = StylesheetInternal.Fills?.Elements<SpreadsheetLib.Fill>().ElementAt(style.FillId);
            if (fill != null)
            {
                differentialFormat.Fill = (SpreadsheetLib.Fill)fill.CloneNode(true);
            }
        }

        if (style.BorderId > 0)
        {
            var border = StylesheetInternal.Borders?.Elements<SpreadsheetLib.Border>().ElementAt(style.BorderId);
            if (border != null)
            {
                differentialFormat.Border = (SpreadsheetLib.Border)border.CloneNode(true);
            }
        }

        var styleElement = GetStyleElement(style);
        if (styleElement.Alignment != null)
        {
            differentialFormat.Alignment = (SpreadsheetLib.Alignment)styleElement.Alignment.CloneNode(true);
        }

        return differentialFormat;
    }

    private ITableInfo AddTableCore(string worksheetName, ICell startCell, ICell endCell, List<string> columnsName, TableCreateOptions? options)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);
        ArgumentNullException.ThrowIfNull(startCell);
        ArgumentNullException.ThrowIfNull(endCell);
        ArgumentNullException.ThrowIfNull(columnsName);

        if (columnsName.Count == 0)
        {
            throw new ArgumentException("Column names list cannot be empty.", nameof(columnsName));
        }

        if (columnsName.Any(string.IsNullOrWhiteSpace))
        {
            throw new ArgumentException("Table column names cannot be null or empty.", nameof(columnsName));
        }

        if (startCell.RowIndex > endCell.RowIndex || startCell.ColumnIndex > endCell.ColumnIndex)
        {
            throw new ArgumentException("Invalid table definition: start cell must be before end cell.");
        }

        var expectedColumnCount = endCell.ColumnIndex - startCell.ColumnIndex + 1;
        if (columnsName.Count != expectedColumnCount)
        {
            throw new ArgumentException("The number of table columns must match the table width.", nameof(columnsName));
        }

        var worksheet = GetWorksheetOrThrow(worksheetName);
        var worksheetPart = worksheet.WorksheetPart;
        var tableIndex = WorkbookPartInternal.WorksheetParts.SelectMany(part => part.TableDefinitionParts).Count() + 1;
        var autoName = $"Table{tableIndex}";
        var tableName = options?.TableName ?? autoName;
        var displayName = options?.DisplayName ?? tableName;

        var styleOptions = options?.Style;
        var tableRef = $"{startCell.CellReference}:{endCell.CellReference}";
        var table = new SpreadsheetLib.Table
        {
            Id = (uint)tableIndex,
            Name = tableName,
            DisplayName = displayName,
            Reference = tableRef,
            TotalsRowShown = false,
            TableColumns = new SpreadsheetLib.TableColumns { Count = Convert.ToUInt32(columnsName.Count) },
            AutoFilter = new SpreadsheetLib.AutoFilter { Reference = tableRef },
            TableStyleInfo = new SpreadsheetLib.TableStyleInfo
            {
                Name = styleOptions?.StyleName ?? "TableStyleMedium2",
                ShowFirstColumn = styleOptions?.ShowFirstColumn ?? false,
                ShowLastColumn = styleOptions?.ShowLastColumn ?? false,
                ShowRowStripes = styleOptions?.ShowBandedRows ?? true,
                ShowColumnStripes = styleOptions?.ShowBandedColumns ?? false
            }
        };

        for (var index = 0; index < columnsName.Count; index++)
        {
            table.TableColumns.Append(new SpreadsheetLib.TableColumn
            {
                Id = (uint)index + 1,
                Name = columnsName[index]
            });
        }

        var tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
        tablePart.Table = table;
        var tableRelationshipId = worksheetPart.GetIdOfPart(tablePart);

        var tableParts = worksheet.WorksheetElement.GetFirstChild<SpreadsheetLib.TableParts>();
        if (tableParts == null)
        {
            tableParts = worksheet.WorksheetElement.AppendChild(new SpreadsheetLib.TableParts());
        }

        tableParts.Append(new SpreadsheetLib.TablePart { Id = tableRelationshipId });
        tableParts.Count = Convert.ToUInt32(tableParts.Elements<SpreadsheetLib.TablePart>().Count());

        return new TableInfo(tableName, displayName, tableRef, columnsName.AsReadOnly(), worksheetName);
    }

    private static TableDefinitionPart? FindTableDefinitionPart(WorksheetPart worksheetPart, string tableName)
    {
        return worksheetPart.TableDefinitionParts
            .FirstOrDefault(part => string.Equals(part.Table?.Name, tableName, StringComparison.OrdinalIgnoreCase));
    }

    private static ITableInfo ToTableInfo(string worksheetName, TableDefinitionPart part)
    {
        var table = part.Table ?? throw new InvalidOperationException("The table definition part does not contain a table.");
        var columnNames = table.TableColumns?
            .Elements<SpreadsheetLib.TableColumn>()
            .OrderBy(col => col.Id?.Value)
            .Select(col => col.Name?.Value ?? string.Empty)
            .ToList() ?? [];

        return new TableInfo(
            table.Name?.Value ?? string.Empty,
            table.DisplayName?.Value ?? string.Empty,
            table.Reference?.Value ?? string.Empty,
            columnNames.AsReadOnly(),
            worksheetName);
    }

    private static SpreadsheetLib.CellFormat GetStyleElement(IStyle style)
    {
        if (style is Style concreteStyle)
        {
            return concreteStyle.ElementInternal;
        }

#pragma warning disable CS0618
        return style.Element;
#pragma warning restore CS0618
    }
}
