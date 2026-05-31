using System.Collections.Generic;
using System.Globalization;
using OfficeDocuments.Excel.Extensions;
using OfficeDocuments.Excel.Interfaces;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.DataClasses;

internal class Cell : Base, ICell
{
    public OpenXml.Cell Element { get; }
    public string CellReference { get; }
    public uint RowIndex => _rowIndex > 0
        ? _rowIndex
        : GetOrCacheCellIndices().rowIndex;
    public uint ColumnIndex => _columnIndex > 0
        ? _columnIndex
        : GetOrCacheCellIndices().columnIndex;

    public string Value
    {
        get => GetStringValue() ?? string.Empty;
        set => SetValue(value);
    }

    private uint _rowIndex;
    private uint _columnIndex;

    private (uint rowIndex, uint columnIndex) GetOrCacheCellIndices()
    {
        if (_rowIndex > 0 && _columnIndex > 0)
        {
            return (_rowIndex, _columnIndex);
        }

        if (!CellExtension.TryParseCellReference(CellReference.AsSpan(), out var rowIndex, out var columnIndex))
        {
            throw new InvalidOperationException($"The cell reference '{CellReference}' is invalid.");
        }

        _rowIndex = rowIndex;
        _columnIndex = columnIndex;
        return (_rowIndex, _columnIndex);
    }

    internal Cell(IWorksheet worksheet, uint column, uint row, IStyle? cellStyle = null)
        : this(worksheet, CellExtension.GetExcelCellReference(column, row), cellStyle)
    {
        _rowIndex = row;
        _columnIndex = column;
    }

    internal Cell(IWorksheet worksheet, string cellReference, IStyle? cellStyle)
        : base(worksheet, cellStyle)
    {
        CellReference = cellReference;
        Element = new OpenXml.Cell
        {
            CellReference = cellReference
        };

        if (Style != null)
        {
            Element.StyleIndex = Convert.ToUInt32(Style.StyleIndex);
        }
    }

    internal Cell(IWorksheet worksheet, OpenXml.Cell element)
        : base(worksheet, element.StyleIndex ?? 0)
    {
        CellReference = element.CellReference?.Value ?? throw new InvalidOperationException("The cell reference is missing.");
        Element = element;
    }

    public void SetValue(object? value)
    {
        if (value == null)
        {
            return;
        }

        switch (Type.GetTypeCode(value.GetType()))
        {
            case TypeCode.Boolean:
                SetValue((bool)value);
                break;
            case TypeCode.Byte:
            case TypeCode.SByte:
            case TypeCode.UInt16:
            case TypeCode.Int16:
            case TypeCode.Int32:
            case TypeCode.UInt32:
            case TypeCode.UInt64:
            case TypeCode.Int64:
                SetNumberValue(value, 1);
                break;
            case TypeCode.Decimal:
            case TypeCode.Double:
            case TypeCode.Single:
                SetNumberValue(value, 4);
                break;
            case TypeCode.DateTime:
                SetValue((DateTime)value);
                break;
            default:
                SetValue(value.ToString() ?? string.Empty);
                break;
        }
    }

    public void SetValue(bool value)
    {
        SetCellValue(value.ToString(CultureInfo.InvariantCulture), OpenXml.CellValues.Boolean);
    }

    public void SetValue(DateTime value)
    {
        if (Style == null || Style.NumberFormatId == 0)
        {
            AddStyle(new Style(OwnerSpreadsheet.StylesheetInternal, 0, 0, 0, 14));
        }

        SetCellValue(value.ToOADate().ToString(CultureInfo.InvariantCulture));
    }

    public void SetValue(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return;
        }

        if (Style == null || Style.NumberFormatId == 0)
        {
            AddStyle(new Style(OwnerSpreadsheet.StylesheetInternal, 0, 0, 0, 49));
        }

        SetCellValue(value, OpenXml.CellValues.String);
    }

    public void SetFormula(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return;
        }

        Element.CellFormula = new OpenXml.CellFormula(formula);
        Element.CellValue = null;
        Element.DataType = null;
    }

    public void SetHyperlink(string target, string? displayText = null)
    {
        OwnerWorksheet.SetCellHyperlink(this, target, displayText);
    }

    public string? GetHyperlink() => OwnerWorksheet.GetCellHyperlink(CellReference);

    public void SetComment(string text, string? author = null)
    {
        OwnerWorksheet.SetCellComment(this, text, author);
    }

    public string? GetComment() => OwnerWorksheet.GetCellComment(CellReference);

    public string? GetFormula() => Element.CellFormula?.Text;

    public int GetFormulaValue()
    {
        var formula = GetFormula();
        if (string.IsNullOrEmpty(formula))
        {
            return -1;
        }

        return formula switch
        {
            var currentFormula when currentFormula.StartsWith("SUM", StringComparison.Ordinal) => FormulaSum(currentFormula),
            var currentFormula when currentFormula.StartsWith("COUNTIF", StringComparison.Ordinal) => CountCellsIf(currentFormula),
            var currentFormula when currentFormula.StartsWith("COUNT", StringComparison.Ordinal) => CountCellsWithValue(currentFormula),
            var currentFormula when currentFormula.StartsWith("MEDIAN", StringComparison.Ordinal) => GetMedian(currentFormula),
            _ => throw new NotImplementedException(),
        };
    }

    public int FormulaSum(string formula)
    {
        var parts = formula.Split('(', ')', ':');
        const string methodName = "SUM";
        var range = parts.Where(part => !string.IsNullOrEmpty(part) && part != methodName).ToArray();
        var (_, fromColumnIndex) = range[0].GetExcelCellIndex();
        var (_, toColumnIndex) = range[1].GetExcelCellIndex();
        var sum = 0;

        for (var columnIndex = fromColumnIndex; columnIndex <= toColumnIndex; columnIndex++)
        {
            var cell = Worksheet.GetCell(columnIndex);
            if (cell == null)
            {
                continue;
            }

            if (cell.HasFormula())
            {
                sum += cell.GetFormulaValue();
                continue;
            }

            sum += cell.TryGetValue(out int value)
                ? value
                : throw new ArgumentException($"Invalid cell '{cell.CellReference}' content.");
        }

        return sum;
    }

    public int CountCellsWithValue(string formula)
    {
        var parts = formula.Split('(', ')', ':');
        const string methodName = "COUNT";
        var range = parts.Where(part => !string.IsNullOrEmpty(part) && part != methodName).ToArray();
        var (_, fromColumnIndex) = range[0].GetExcelCellIndex();
        var (_, toColumnIndex) = range[1].GetExcelCellIndex();
        var sum = 0;

        for (var columnIndex = fromColumnIndex; columnIndex <= toColumnIndex; columnIndex++)
        {
            if (Worksheet.GetCell(columnIndex)?.HasValue() == true)
            {
                sum++;
            }
        }

        return sum;
    }

    public int CountCellsIf(string formula)
    {
        var parts = formula.Split('(', ')', ':', ',');
        const string methodName = "COUNTIF";
        var range = parts.Where(part => !string.IsNullOrEmpty(part) && part != methodName).ToArray();
        var (_, fromColumnIndex) = range[0].GetExcelCellIndex();
        var (_, toColumnIndex) = range[1].GetExcelCellIndex();
        var argument = range[2];
        var argumentValue = argument.StartsWith("\"", StringComparison.Ordinal) && argument.EndsWith("\"", StringComparison.Ordinal)
            ? argument.Trim('"')
            : ResolveArgumentValue(argument);
        var sum = 0;

        for (var columnIndex = fromColumnIndex; columnIndex <= toColumnIndex; columnIndex++)
        {
            var cell = Worksheet.GetCell(columnIndex);
            if (cell?.HasValue() == true && cell.Value == argumentValue)
            {
                sum++;
            }
            else if (cell?.HasValue() != true && argumentValue == string.Empty)
            {
                sum++;
            }
        }

        return sum;
    }

    public int GetMedian(string formula)
    {
        var parts = formula.Split('(', ')', ':');
        const string methodName = "MEDIAN";
        var range = parts.Where(part => !string.IsNullOrEmpty(part) && part != methodName).ToArray();
        var (_, fromColumnIndex) = range[0].GetExcelCellIndex();
        var (_, toColumnIndex) = range[1].GetExcelCellIndex();
        var values = new List<int>();

        for (var columnIndex = fromColumnIndex; columnIndex <= toColumnIndex; columnIndex++)
        {
            var cell = Worksheet.GetCell(columnIndex);
            if (cell == null)
            {
                continue;
            }

            if (cell.HasFormula())
            {
                values.Add(cell.GetFormulaValue());
            }
            else if (cell.TryGetValue(out int value))
            {
                values.Add(value);
            }
        }

        return Median(values.ToArray());
    }

    public static int Median(int[] data)
    {
        Array.Sort(data);
        return data.Length % 2 == 0
            ? (data[data.Length / 2 - 1] + data[data.Length / 2]) / 2
            : data[data.Length / 2];
    }

    public string? GetStringValue()
    {
        if (HasFormula())
        {
            throw new InvalidOperationException($"Cell '{CellReference}': Cannot get value of formula");
        }

        if (Element.InlineString?.Text != null)
        {
            return Element.InlineString.Text.Text;
        }

        var value = Element.CellValue?.Text;
        if (!string.IsNullOrEmpty(value) && Element.DataType?.Value == OpenXml.CellValues.SharedString && int.TryParse(value.Trim(), out var stringId))
        {
            var item = GetSharedStringItemById(stringId);
            if (item.Text != null)
            {
                return item.Text.Text;
            }

            if (item.InnerText != null)
            {
                return item.InnerText;
            }
        }

        return value;
    }

    public bool GetBoolValue() => GetValue(bool.Parse);

    public int GetIntValue() => GetValue(int.Parse);

    public long GetLongValue() => GetValue(long.Parse);

    public double GetDoubleValue() => GetInvariantValue(double.Parse);

    public decimal GetDecimalValue() => GetInvariantValue(decimal.Parse);

    public DateTime GetDateValue(string? format = null)
    {
        var cellValue = GetStringValue();
        if (string.IsNullOrEmpty(cellValue))
        {
            throw new InvalidOperationException($"Cell '{CellReference}' does not contain a value.");
        }

        return format == null
            ? DateTime.FromOADate(double.Parse(cellValue, CultureInfo.InvariantCulture))
            : DateTime.ParseExact(cellValue, format, CultureInfo.InvariantCulture);
    }

    public bool TryGetValue(out bool value) => bool.TryParse(GetStringValue(), out value);

    public bool TryGetValue(out int value) => int.TryParse(GetStringValue(), out value);

    public bool TryGetValue(out long value) => long.TryParse(GetStringValue(), out value);

    public bool TryGetValue(out double value) => double.TryParse(GetStringValue(), NumberStyles.Any, CultureInfo.InvariantCulture, out value);

    public bool TryGetValue(out decimal value) => decimal.TryParse(GetStringValue(), NumberStyles.Any, CultureInfo.InvariantCulture, out value);

    public bool TryGetValue(out string value)
    {
        value = string.Empty;
        if (HasFormula())
        {
            return false;
        }

        var stringValue = GetStringValue();
        if (stringValue == null)
        {
            return false;
        }

        value = stringValue;
        return true;
    }

    public bool TryGetValue(out DateTime value, string? format = null)
    {
        value = DateTime.MinValue;
        if (HasFormula())
        {
            return false;
        }

        var stringValue = GetStringValue();
        if (string.IsNullOrEmpty(stringValue))
        {
            return false;
        }

        if (format == null && double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var oaValue))
        {
            value = DateTime.FromOADate(oaValue);
            return true;
        }

        return DateTime.TryParseExact(stringValue, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out value);
    }

    public bool HasValue() => !string.IsNullOrEmpty(Element.CellValue?.Text) || Element.InlineString != null;

    public bool HasFormula() => !string.IsNullOrEmpty(Element.CellFormula?.Text);

    public override IStyle? AddStyle(params IStyle?[] styles)
    {
        foreach (var style in styles.Where(currentStyle => currentStyle != null))
        {
            Style = Style?.CreateMergedStyle(style) ?? style;
        }

        if (Style != null)
        {
            Element.StyleIndex = Convert.ToUInt32(Style.StyleIndex);
        }

        return Style;
    }

    internal OpenXml.Cell? CloneElement() => (OpenXml.Cell)Element.CloneNode(true);

    internal void ReplaceFrom(OpenXml.Cell? sourceCell)
    {
        Element.RemoveAllChildren();
        Element.CellFormula = null;
        Element.CellValue = null;
        Element.InlineString = null;
        Element.DataType = null;
        Element.StyleIndex = null;

        if (sourceCell != null)
        {
            foreach (var child in sourceCell.ChildElements)
            {
                Element.Append(child.CloneNode(true));
            }

            Element.DataType = sourceCell.DataType?.Value;
            Element.StyleIndex = sourceCell.StyleIndex?.Value;
        }

        Element.CellReference = CellReference;
        Style = Element.StyleIndex == null ? null : new Style(OwnerSpreadsheet.StylesheetInternal, Convert.ToInt32(Element.StyleIndex.Value));
    }

    private void SetNumberValue(object value, int numberFormatId)
    {
        if (Style == null || Style.NumberFormatId == 0)
        {
            AddStyle(new Style(OwnerSpreadsheet.StylesheetInternal, numberFormatId: numberFormatId));
        }

        SetCellValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, OpenXml.CellValues.Number);
    }

    private string ResolveArgumentValue(string argument)
    {
        var cell = Worksheet.GetCellByReference(argument);
        if (cell == null)
        {
            return argument;
        }

        return cell.HasFormula()
            ? cell.GetFormulaValue().ToString(CultureInfo.InvariantCulture)
            : cell.HasValue()
                ? cell.Value
                : string.Empty;
    }

    private void SetCellValue(string value, OpenXml.CellValues? dataType = null)
    {
        Element.CellFormula = null;
        Element.CellValue = new OpenXml.CellValue(value);
        if (dataType != null && dataType != OpenXml.CellValues.Error)
        {
            Element.DataType = dataType;
        }
        else
        {
            Element.DataType = null;
        }
    }

    private OpenXml.SharedStringItem GetSharedStringItemById(int id)
    {
        var sharedStringTable = OwnerSpreadsheet.WorkbookPartInternal.SharedStringTablePart?.SharedStringTable
            ?? throw new InvalidOperationException("The workbook does not contain a shared string table.");

        return sharedStringTable.Elements<OpenXml.SharedStringItem>().ElementAt(id);
    }

    private T GetValue<T>(Func<string, T> parse) where T : IConvertible => parse(GetRequiredStringValue());

    private T GetInvariantValue<T>(Func<string, IFormatProvider, T> parse) where T : IConvertible => parse(GetRequiredStringValue(), CultureInfo.InvariantCulture);

    private string GetRequiredStringValue() => GetStringValue() ?? throw new InvalidOperationException($"Cell '{CellReference}' does not contain a value.");
}
