using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Office2016.Presentation.Command;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocumentsApi.Excel.Extensions;
using OfficeDocumentsApi.Excel.Interfaces;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocumentsApi.Excel.DataClasses
{
    internal class Cell : Base, ICell
    {
        private delegate T ParseDelegate<out T>(string s, IFormatProvider provider);

        public OpenXml.Cell Element { get; }

        public string CellReference { get; }
        public uint RowIndex => _rowIndex > 0
            ? _rowIndex
            : _rowIndex = uint.Parse(new string(CellReference.Where(char.IsDigit).ToArray()));

        public uint ColumnIndex => _columnIndex > 0
            ? _columnIndex
            : _columnIndex = new string(CellReference.Where(char.IsLetter).ToArray()).GetExcelColumnIndex();
        // : _columnIndex = GetExcelColumnIndex(new string(CellReference.Where(char.IsLetter).ToArray()));

        public string Value
        {
            get => GetStringValue();
            set => SetValue(value);
        }

        private uint _rowIndex;
        private uint _columnIndex;

        internal Cell(IWorksheet worksheet, uint column, uint row, IStyle? cellStyle = null)
            : this(worksheet, GetExcelColumnName(column) + row, cellStyle)
        {
            _rowIndex = row;
            _columnIndex = column;
        }
        internal Cell(IWorksheet worksheet, string cellReference, IStyle cellStyle)
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
            CellReference = element.CellReference;

            Element = element;
        }

        #region Set value/formula

        public void SetValue(object value)
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
                    SetNumberValue(value, 1); // value 1 as number format "0"
                    break;
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    SetNumberValue(value, 4); // value 4 as number format "#,##0.00"
                    break;

                case TypeCode.DateTime:
                    SetValue((DateTime)value);
                    break;

                case TypeCode.String:
                default:
                    SetValue(value.ToString());
                    break;
            }
        }
        public void SetValue(bool value)
        {
            SetCellValue(value.ToString(CultureInfo.InvariantCulture), OpenXml.CellValues.Boolean);
        }

        private void SetNumberValue<TNumber>(TNumber value, int numberFormatId) where TNumber : class
        {
            if (Style == null || Style.NumberFormatId == 0)
            {
                var s = new Style(Worksheet.Spreadsheet.WorkbookStylesPart.Stylesheet, numberFormatId: numberFormatId); // "0"
                AddStyle(s);
            }

            SetCellValue(((IConvertible)value).ToString(CultureInfo.InvariantCulture), OpenXml.CellValues.Number);
        }

        public void SetValue(DateTime value)
        {
            if (Style == null || Style.NumberFormatId == 0)
            {
                var s = new Style(Worksheet.Spreadsheet.WorkbookStylesPart.Stylesheet, 0, 0, 0, 14); // "d/m/yyyy"
                AddStyle(s);
            }

            // cell with date needs Number format for DateTime, not DataType
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
                var s = new Style(Worksheet.Spreadsheet.WorkbookStylesPart.Stylesheet, 0, 0, 0, 49); // "@"
                AddStyle(s);
            }

            SetCellValue(value, OpenXml.CellValues.String);
        }

        public void SetFormula(string formula)
        {
            if (string.IsNullOrEmpty(formula))
            {
                return;
            }
            Element.CellFormula = new OpenXml.CellFormula(formula);
        }
        #endregion

        #region Get value/formula
        public string? GetFormula()
        {
            return Element.CellFormula?.Text;
        }

        public int GetFormulaValue()
        {
            if (!HasFormula())
                return -1;

            return GetFormula() switch
            {
                var f when f.StartsWith("SUM") && false => FormulaSum(f), // if any cell is double
                var f when f.StartsWith("SUM") => FormulaSum(f),
                var f when f.StartsWith("COUNTIF") => CountCellsIf(f),
                var f when f.StartsWith("COUNT") => CountCellsWithValue(f),
                var f when f.StartsWith("MEDIAN") => GetMedian(f),
                _ => throw new NotImplementedException(),
            };
        }

        public int FormulaSum(string formula)
        {
            // Split formula to cell names in string array
            var subs = formula.Split('(', ')', ':');
            var sum = 0;
            const string methodName = "SUM";
            var range = subs.Where(x => !string.IsNullOrEmpty(x) && x != methodName).ToArray();

            var (_, fromColumnIndex) = range[0].GetExcelCellIndex();
            var (_, toColumnIndex) = range[1].GetExcelCellIndex();

            for (var i = fromColumnIndex; i <= toColumnIndex; i++)
            {
                var cell = Worksheet.GetCell(i);
                if (cell == null)
                {
                    continue;
                }

                if (cell.HasFormula())
                {
                    sum += cell.GetFormulaValue();
                    continue;
                }

                // TODO get cell; if cell is formula => get formula value       DONE
                // TODO get cell; if cell is text => throw exception            DONE
                // TODO get cell; if any of cell is double => return double

                sum += 1 switch
                {
                    _ when cell.TryGetValue(out int val) => val,
                    // when cell.TryGetValue(out double val) => val,
                    _ => throw new ArgumentException($"Invalid cell '{cell.CellReference}' content."),
                };
            }

            return sum;
        }

        public int CountCellsWithValue(string formula)
        {
            // Split formula to cell names in string array
            string[] subs = formula.Split('(', ')', ':');
            var sum = 0;

            const string methodName = "COUNT";
            var range = subs.Where(x => !string.IsNullOrEmpty(x) && x != methodName).ToArray();

            var (_, fromColumnIndex) = range[0].GetExcelCellIndex();
            var (_, toColumnIndex) = range[1].GetExcelCellIndex();

            for (var i = fromColumnIndex; i <= toColumnIndex; i++)
            {
                var cell = Worksheet.GetCell(i);

                if (cell.HasValue())
                {
                    sum++;
                }
            }

            return sum;
        }

        public int CountCellsIf(string formula)
        {
            // Split formula to cell names in string array
            string[] subs = formula.Split('(', ')', ':', ',');
            var sum = 0;

            const string methodName = "COUNTIF";
            var range = subs.Where(x => !string.IsNullOrEmpty(x) && x != methodName).ToArray();

            var (_, fromColumnIndex) = range[0].GetExcelCellIndex();
            var (_, toColumnIndex) = range[1].GetExcelCellIndex();
            var argument = range[2];
            var argumentValue = string.Empty;

            if (argument.StartsWith("\"") && argument.EndsWith("\""))
            {
                argumentValue = argument.Trim('\"');
            }
            else
            {
                var cell = Worksheet.GetCellByReference(argument);
                if (cell != null)
                {
                    argumentValue = cell.HasFormula()
                        ? cell.GetFormulaValue().ToString()
                        : cell.HasValue()
                            ? cell.Value
                            : string.Empty;
                }
                else
                {
                    argumentValue = argument;
                }
            }

            for (var i = fromColumnIndex; i <= toColumnIndex; i++)
            {
                var cell = Worksheet.GetCell(i);

                if (cell.HasValue() && cell.Value == argumentValue)
                {
                    sum++;
                }
                else if (!cell.HasValue() && argumentValue == string.Empty)
                {
                    sum++;
                }
            }

            return sum;
        }

        public int GetMedian(string formula)
        {
            string[] subs = formula.Split('(', ')', ':');

            const string methodName = "MEDIAN";
            var range = subs.Where(x => !string.IsNullOrEmpty(x) && x != methodName).ToArray();


            var (_, fromColumnIndex) = range[0].GetExcelCellIndex();
            var (_, toColumnIndex) = range[1].GetExcelCellIndex();

            // Check if value is a number
            // If true, add coordinates to list
            var columns = new List<int>();

            for (var i = fromColumnIndex; i <= toColumnIndex; i++)
            {
                var cell = Worksheet.GetCell(i);
                if (cell != null)
                {
                    int value;
                    switch (1)
                    {
                        case 1 when cell.HasFormula():
                            value = cell.GetFormulaValue();
                            break;
                        case 1 when cell.TryGetValue(out int v):
                            value = v;
                            break;
                        default:
                            continue;
                    }

                    columns.Add(value);
                }
            }
            /*

            // Calculate median
            var middleCell = columns.Count / 2;
            var median = 0;

            if (columns.Count % 2 == 0)
            {
                var cell1 = Worksheet.GetCell(columns[middleCell - 1]);
                var cell2 = Worksheet.GetCell(columns[middleCell - 2]);

                median = cell1.GetIntValue() + cell2.GetIntValue() / 2;
            }
            else
            {
                median = Worksheet.GetCell(columns[middleCell - 1]).GetIntValue();
            }*/

            return Median(columns.ToArray());
        }

        public static int Median(int[] data)
        {
            Array.Sort(data);

            if (data.Length % 2 == 0)
                return (data[data.Length / 2 - 1] + data[data.Length / 2]) / 2;
            else
                return data[data.Length / 2];
        }

        public string? GetStringValue()
        {
            if (HasFormula())
            {
                throw new InvalidOperationException($"Cell '{CellReference}': Cannot get value of formula");
            }

            var value = Element.CellValue?.Text;

            if (!string.IsNullOrEmpty(value) && Element.DataType?.Value == OpenXml.CellValues.SharedString)
            {
                var stringId = -1;

                if (int.TryParse(value.Trim(), out stringId))
                {
                    var item = GetSharedStringItemById(stringId);

                    if (item.Text != null)
                    {
                        value = item.Text.Text;
                    }
                }
            }
            return value;
        }

        public bool GetBoolValue()
        {
            return GetValue(bool.Parse);
        }

        public int GetIntValue()
        {
            return GetValue(int.Parse);
        }

        public long GetLongValue()
        {
            return GetValue(long.Parse);
        }

        public double GetDoubleValue()
        {
            return GetInvariantValue(double.Parse);
        }

        public decimal GetDecimalValue()
        {
            return GetInvariantValue(decimal.Parse);
        }

        public DateTime GetDateValue(string? format = null)
        {
            var cellValue = GetStringValue();

            return format == null
                    ? DateTime.FromOADate(double.Parse(cellValue, CultureInfo.InvariantCulture))
                    : DateTime.ParseExact(cellValue, format, CultureInfo.InvariantCulture);
        }

        public bool TryGetValue(out bool value)
        {
            return bool.TryParse(GetStringValue(), out value);
        }

        public bool TryGetValue(out int value)
        {
            return int.TryParse(GetStringValue(), out value);
        }

        public bool TryGetValue(out long value)
        {
            return long.TryParse(GetStringValue(), out value);
        }

        public bool TryGetValue(out double value)
        {
            return double.TryParse(GetStringValue(), NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        public bool TryGetValue(out decimal value)
        {
            return decimal.TryParse(GetStringValue(), NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        public bool TryGetValue(out string value)
        {
            value = string.Empty;
            try
            {
                value = GetStringValue();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool TryGetValue(out DateTime value, string? format = null)
        {
            value = DateTime.MinValue;
            try
            {
                value = GetDateValue(format);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool HasValue()
        {
            return !string.IsNullOrEmpty(Element.CellValue?.Text);
        }

        public bool HasFormula()
        {
            return !string.IsNullOrEmpty(Element.CellFormula?.Text);
        }
        #endregion

        public override IStyle? AddStyle(params IStyle[] styles)
        {
            foreach (var style in styles.Where(s => s != null))
            {
                Style = Style?.CreateMergedStyle(style) ?? style;
            }

            if (Style != null && Element != null)
            {
                Element.StyleIndex = Convert.ToUInt32(Style.StyleIndex);
            }

            return Style;
        }

        private static string GetExcelColumnName(uint columnIndex)
        {
            var dividend = columnIndex; // A column is column number 1
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(64 + 1 + modulo) + columnName;
                dividend = (uint)((dividend - modulo) / 26);
            }

            return columnName;
        }
        /*
        private static uint GetExcelColumnIndex(string columnName)
        {
            return (uint)columnName
                .ToUpper()
                .Aggregate(0, (column, letter) => 26 * column + letter - 'A' + 1);
        }*/

        private void SetCellValue(string value, OpenXml.CellValues dataType = OpenXml.CellValues.Error)
        {
            Element.CellValue = new OpenXml.CellValue(value);
            if (dataType != OpenXml.CellValues.Error)
            {
                Element.DataType = dataType;
            }
        }
        private OpenXml.SharedStringItem GetSharedStringItemById(int id)
        {
            return Worksheet.Spreadsheet.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<OpenXml.SharedStringItem>().ElementAt(id);
        }

        private T GetValue<T>(Func<string, T> parse) where T : IConvertible
        {
            return parse(GetStringValue());
        }
        private T GetInvariantValue<T>(Func<string, IFormatProvider, T> parse) where T : IConvertible
        {
            return parse(GetStringValue(), CultureInfo.InvariantCulture);
        }
    }
}