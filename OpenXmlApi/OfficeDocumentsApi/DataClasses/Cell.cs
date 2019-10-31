using System;
using System.Globalization;
using System.Linq;
using OfficeDocumentsApi.Interfaces;
using OfficeDocumentsApi.Styles;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocumentsApi.DataClasses
{
    internal class Cell : Base, ICell
    {
        private delegate T ParseDelegate<out T>(string s, IFormatProvider provider);

        public OpenXml.Cell Element { get; }

        public string CellReference { get; }
        public uint RowIndex => 
            rowIndex > 0
            ? rowIndex
            : rowIndex = uint.Parse(new string(CellReference.Where(char.IsDigit).ToArray()));

        public uint ColumnIndex => columnIndex > 0
            ? columnIndex
            : columnIndex = GetExcelColumnIndex(new string(CellReference.Where(char.IsLetter).ToArray()));

        public string Value
        {
            get => GetStringValue();
            set => SetValue(value);
        }

        private uint rowIndex;
        private uint columnIndex;

        internal Cell(IWorksheet worksheet, uint column, uint row, IStyle cellStyle = null)
            : this(worksheet, GetExcelColumnName(column) + row, cellStyle)
        {
            rowIndex = row;
            columnIndex = column;
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
        
        private void SetNumberValue<TNumber>(TNumber value, int numberFormatId) where  TNumber : class
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
        public string GetFormula() => Element.CellFormula?.Text;

        public string GetStringValue()
        {
            if (HasFormula())
            {
                throw new InvalidOperationException($"Cell '{CellReference}': Cannot get value of formula");
            }

            var value = Element.CellValue?.Text;

            if (!string.IsNullOrEmpty(value) && Element.DataType?.Value == OpenXml.CellValues.SharedString)
            {
                if (int.TryParse(value.Trim(), out var stringId))
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

        public bool GetBoolValue() => GetValue(bool.Parse);

        public int GetIntValue() => GetValue(int.Parse);

        public long GetLongValue() => GetValue(long.Parse);

        public double GetDoubleValue() => GetInvariantValue(double.Parse);

        public decimal GetDecimalValue() => GetInvariantValue(decimal.Parse);

        public DateTime GetDateValue(string format = null)
        {
            var cellValue = GetStringValue();

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

        public bool TryGetValue(out DateTime value, string format = null)
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

        public bool HasValue() => !string.IsNullOrEmpty(Element.CellValue?.Text);

        public bool HasFormula() => !string.IsNullOrEmpty(Element.CellFormula?.Text);

        #endregion

        public override IStyle AddStyle(params IStyle[] styles)
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
                dividend = (dividend - modulo) / 26u;
            }

            return columnName;
        }

        private static uint GetExcelColumnIndex(string columnName) =>
            (uint)columnName.ToUpper().Aggregate(0, (column, letter) => 26 * column + letter - 'A' + 1);

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