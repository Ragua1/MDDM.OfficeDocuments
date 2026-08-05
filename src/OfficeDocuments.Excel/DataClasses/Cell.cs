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
        SetCellValue(value ? "1" : "0", OpenXml.CellValues.Boolean);
    }

    public void SetValue(DateTime value)
    {
        // ExcelSerialDate, not ToOADate: the two disagree by a day before March 1900. See that
        // type for why, and why using ToOADate on both sides hides the problem rather than solving it.
        var serial = ExcelSerialDate.ToSerial(value);

        if (Style == null || Style.NumberFormatId == 0)
        {
            AddStyle(new Style(OwnerSpreadsheet.StylesheetInternal, 0, 0, 0, 14));
        }

        SetCellValue(serial.ToString(CultureInfo.InvariantCulture));
    }

    public void SetValue(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return;
        }

        XmlText.EnsureRepresentable(value, nameof(value), $"The value for cell '{CellReference}'");

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

    public double GetFormulaValue()
    {
        var formula = GetFormula();
        if (string.IsNullOrEmpty(formula))
        {
            throw new InvalidOperationException($"Cell '{CellReference}' does not contain a formula.");
        }

        var functionName = GetFunctionName(formula);
        return functionName switch
        {
            "SUM" => FormulaSum(formula),
            "COUNTIF" => CountCellsIf(formula),
            "COUNT" => CountCellsWithValue(formula),
            "MEDIAN" => GetMedian(formula),
            _ => throw new NotSupportedException($"Formula function '{functionName}' is not supported by the built-in evaluator."),
        };
    }

    private static string GetFunctionName(string formula)
    {
        var parenthesisIndex = formula.IndexOf('(');
        var name = parenthesisIndex < 0 ? formula : formula[..parenthesisIndex];
        return name.Trim().ToUpperInvariant();
    }

    private (uint fromColumn, uint fromRow, uint toColumn, uint toRow) GetFormulaRange(string formula)
    {
        var open = formula.IndexOf('(');
        var close = formula.LastIndexOf(')');
        if (open < 0 || close < open)
        {
            throw new ArgumentException($"Malformed formula '{formula}'.");
        }

        var rangeToken = formula[(open + 1)..close].Split(',')[0].Trim();
        if (!rangeToken.TryGetExcelRange(out var coordinates))
        {
            throw new ArgumentException($"Invalid range '{rangeToken}' in formula '{formula}'.");
        }

        return coordinates;
    }

    private IEnumerable<ICell> GetFormulaRangeCells(string formula)
    {
        var (fromColumn, fromRow, toColumn, toRow) = GetFormulaRange(formula);
        for (var row = fromRow; row <= toRow; row++)
        {
            for (var column = fromColumn; column <= toColumn; column++)
            {
                var cell = Worksheet.GetCell(column, row);
                if (cell != null)
                {
                    yield return cell;
                }
            }
        }
    }

    private double FormulaSum(string formula)
    {
        var sum = 0d;
        foreach (var cell in GetFormulaRangeCells(formula))
        {
            if (cell.HasFormula())
            {
                sum += cell.GetFormulaValue();
                continue;
            }

            sum += cell.TryGetValue(out double value)
                ? value
                : throw new ArgumentException($"Invalid cell '{cell.CellReference}' content.");
        }

        return sum;
    }

    private double CountCellsWithValue(string formula)
    {
        var count = 0d;
        foreach (var cell in GetFormulaRangeCells(formula))
        {
            if (cell.HasValue())
            {
                count++;
            }
        }

        return count;
    }

    private double CountCellsIf(string formula)
    {
        var open = formula.IndexOf('(');
        var close = formula.LastIndexOf(')');
        var arguments = formula[(open + 1)..close].Split(',');
        if (arguments.Length < 2)
        {
            throw new ArgumentException($"COUNTIF requires a range and a criterion in formula '{formula}'.");
        }

        var argument = arguments[1].Trim();
        var argumentValue = argument.StartsWith('"') && argument.EndsWith('"')
            ? argument.Trim('"')
            : ResolveArgumentValue(argument);

        var (fromColumn, fromRow, toColumn, toRow) = GetFormulaRange(formula);
        var count = 0d;
        for (var row = fromRow; row <= toRow; row++)
        {
            for (var column = fromColumn; column <= toColumn; column++)
            {
                var cell = Worksheet.GetCell(column, row);
                if (cell?.HasValue() == true && cell.Value == argumentValue)
                {
                    count++;
                }
                else if (cell?.HasValue() != true && argumentValue == string.Empty)
                {
                    count++;
                }
            }
        }

        return count;
    }

    private double GetMedian(string formula)
    {
        var values = new List<double>();
        foreach (var cell in GetFormulaRangeCells(formula))
        {
            if (cell.HasFormula())
            {
                values.Add(cell.GetFormulaValue());
            }
            else if (cell.TryGetValue(out double value))
            {
                values.Add(value);
            }
        }

        return Median(values);
    }

    private static double Median(IReadOnlyList<double> data)
    {
        if (data.Count == 0)
        {
            throw new InvalidOperationException("MEDIAN requires at least one numeric value in the range.");
        }

        var sorted = data.OrderBy(value => value).ToArray();
        var middle = sorted.Length / 2;
        return sorted.Length % 2 == 0
            ? (sorted[middle - 1] + sorted[middle]) / 2d
            : sorted[middle];
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

    public bool GetBoolValue() => GetValue(ParseBoolean);

    public int GetIntValue() => GetInvariantValue(int.Parse);

    public long GetLongValue() => GetInvariantValue(long.Parse);

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

    public bool TryGetValue(out bool value)
    {
        var raw = GetStringValue();
        switch (raw)
        {
            case "1":
                value = true;
                return true;
            case "0":
                value = false;
                return true;
            default:
                return bool.TryParse(raw, out value);
        }
    }

    public bool TryGetValue(out int value) => int.TryParse(GetStringValue(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

    public bool TryGetValue(out long value) => long.TryParse(GetStringValue(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

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

        if (format == null && double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var serial))
        {
            value = ExcelSerialDate.FromSerial(serial);
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
        EnsureFinite(value);

        if (Style == null || Style.NumberFormatId == 0)
        {
            AddStyle(new Style(OwnerSpreadsheet.StylesheetInternal, numberFormatId: numberFormatId));
        }

        SetCellValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, OpenXml.CellValues.Number);
    }

    /// <summary>
    /// SpreadsheetML has no spelling for NaN or an infinity — a numeric cell holds a decimal
    /// literal and nothing else. Left alone these reach the file verbatim, as
    /// <c>&lt;v&gt;NaN&lt;/v&gt;</c> inside a cell marked <c>t="n"</c>, and Excel reports the
    /// workbook as corrupt.
    /// <para>
    /// Nothing else in the suite catches this. The schema validator will not: <c>v</c> is declared
    /// as a string and the numeric constraint comes from the cell's <c>t</c> attribute, which is a
    /// semantic rule rather than a grammatical one. A round trip will not either, because
    /// <c>double.Parse</c> reads "NaN" straight back. The value has to be refused at the door.
    /// </para>
    /// </summary>
    private void EnsureFinite(object value)
    {
        var isFinite = value switch
        {
            double number => double.IsFinite(number),
            float number => float.IsFinite(number),
            _ => true
        };

        if (!isFinite)
        {
            throw new ArgumentException(
                $"Cell '{CellReference}' cannot hold '{value}'. SpreadsheetML numeric cells have no representation "
                + "for NaN or infinity; write the value as text, or as an Excel error value, if the workbook has to carry it.",
                nameof(value));
        }
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

    private static bool ParseBoolean(string value) => value switch
    {
        "1" => true,
        "0" => false,
        _ => bool.Parse(value),
    };
}
