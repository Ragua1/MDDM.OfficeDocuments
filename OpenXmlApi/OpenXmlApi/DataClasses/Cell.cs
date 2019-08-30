using System;
using System.Globalization;
using System.Linq;
using OpenXmlApi.Interfaces;
using OpenXmlApi.Styles;
using OpenXmlSs = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlApi.DataClasses
{
    internal class Cell : Base, ICell
    {
        public OpenXmlSs.Cell Element { get; }

        public string CellReference { get; }
        public uint RowIndex => this.rowIndex > 0
            ? this.rowIndex
            : this.rowIndex = uint.Parse(new string(this.CellReference.Where(char.IsDigit).ToArray()));

        public uint ColumnIndex => this.columnIndex > 0
            ? this.columnIndex
            : this.columnIndex = GetExcelColumnIndex(new string(this.CellReference.Where(char.IsLetter).ToArray()));

        public string Value
        {
            get => this.GetStringValue();
            set => this.SetValue(value);
        }

        private uint rowIndex = 0;
        private uint columnIndex = 0;

        internal Cell(IWorksheet worksheet, uint column, uint row, IStyle cellStyle = null)
            : this(worksheet, GetExcelColumnName(column) + row, cellStyle)
        {
            this.rowIndex = row;
            this.columnIndex = column;
        }
        internal Cell(IWorksheet worksheet, string cellReference, IStyle cellStyle)
            : base(worksheet, cellStyle)
        {
            this.CellReference = cellReference;

            this.Element = new OpenXmlSs.Cell
            {
                CellReference = cellReference
            };

            if (this.Style != null)
            {
                this.Element.StyleIndex = Convert.ToUInt32(this.Style.StyleIndex);
            }
        }
        internal Cell(IWorksheet worksheet, OpenXmlSs.Cell element)
            : base(worksheet, element.StyleIndex ?? 0)
        {
            this.CellReference = element.CellReference;

            this.Element = element;
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
                    SetValue(Convert.ToInt32(value));
                    break;

                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int64:
                    SetValue(Convert.ToInt64(value));
                    break;

                case TypeCode.Decimal:
                    SetValue(Convert.ToDecimal(value));
                    break;

                case TypeCode.Double:
                case TypeCode.Single:
                    SetValue(Convert.ToDouble(value));
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
            SetCellValue(value.ToString(CultureInfo.InvariantCulture), OpenXmlSs.CellValues.Boolean);
        }
        public void SetValue(int value)
        {
            if (this.Style == null || this.Style.NumberFormatId == 0)
            {
                var s = new Style(this.Worksheet.Spreadsheet.WorkbookStylesPart.Stylesheet, 0, 0, 0, 1); // "0"
                AddStyle(s);
            }

            SetCellValue(value.ToString(CultureInfo.InvariantCulture), OpenXmlSs.CellValues.Number);
        }
        public void SetValue(long value)
        {
            if (this.Style == null || this.Style.NumberFormatId == 0)
            {
                var s = new Style(this.Worksheet.Spreadsheet.WorkbookStylesPart.Stylesheet, 0, 0, 0, 1); // "0"
                AddStyle(s);
            }

            SetCellValue(value.ToString(CultureInfo.InvariantCulture), OpenXmlSs.CellValues.Number);
        }
        public void SetValue(double value)
        {
            if (this.Style == null || this.Style.NumberFormatId == 0)
            {
                var s = new Style(this.Worksheet.Spreadsheet.WorkbookStylesPart.Stylesheet, 0, 0, 0, 4); // "#,##0.00"
                AddStyle(s);
            }

            SetCellValue(value.ToString(CultureInfo.InvariantCulture), OpenXmlSs.CellValues.Number);
        }
        public void SetValue(decimal value)
        {
            if (this.Style == null || this.Style.NumberFormatId == 0)
            {
                var s = new Style(this.Worksheet.Spreadsheet.WorkbookStylesPart.Stylesheet, 0, 0, 0, 4); // "#,##0.00"
                AddStyle(s);
            }

            SetCellValue(value.ToString(CultureInfo.InvariantCulture), OpenXmlSs.CellValues.Number);
        }
        public void SetValue(DateTime value)
        {
            if (this.Style == null || this.Style.NumberFormatId == 0)
            {
                var s = new Style(this.Worksheet.Spreadsheet.WorkbookStylesPart.Stylesheet, 0, 0, 0, 14); // "d/m/yyyy"
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

            if (this.Style == null || this.Style.NumberFormatId == 0)
            {
                var s = new Style(this.Worksheet.Spreadsheet.WorkbookStylesPart.Stylesheet, 0, 0, 0, 49); // "@"
                AddStyle(s);
            }

            SetCellValue(value, OpenXmlSs.CellValues.String);
        }

        public void SetFormula(string formula)
        {
            if (string.IsNullOrEmpty(formula))
            {
                return;
            }
            this.Element.CellFormula = new OpenXmlSs.CellFormula(formula);
        }
        #endregion

        #region Get value/formula
        public string GetFormula()
        {
            return this.Element.CellFormula?.Text;
        }

        public string GetStringValue()
        {
            if (HasFormula())
            {
                throw new InvalidOperationException($"Cell '{this.CellReference}': Cannot get value of formula");
            }

            var value = this.Element.CellValue?.Text;

            if (!string.IsNullOrEmpty(value) && this.Element.DataType?.Value == OpenXmlSs.CellValues.SharedString)
            {
                int stringId = -1;

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

        public DateTime GetDateValue(string format = null)
        {
            var cellValue = GetStringValue();
            DateTime value;

            try
            {
                value = format == null
                    ? DateTime.FromOADate(double.Parse(cellValue, CultureInfo.InvariantCulture))
                    : DateTime.ParseExact(cellValue, format, CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                throw;// new ApplicationServerException(MethodResult.IncorrectFormat);
            }

            return value;
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

        public bool HasValue()
        {
            return !string.IsNullOrEmpty(this.Element.CellValue?.Text);
        }

        public bool HasFormula()
        {
            return !string.IsNullOrEmpty(this.Element.CellFormula?.Text);
        }
        #endregion

        public override IStyle AddStyle(params IStyle[] styles)
        {
            foreach (var style in styles.Where(s => s != null))
            {
                this.Style = this.Style?.CreateMergedStyle(style) ?? style;
            }

            if (this.Style != null && this.Element != null)
            {
                this.Element.StyleIndex = Convert.ToUInt32(this.Style.StyleIndex);
            }

            return this.Style;
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

        private static uint GetExcelColumnIndex(string columnName)
        {
            return (uint)columnName.ToUpper().
                Aggregate(0, (column, letter) => 26 * column + letter - 'A' + 1);
        }

        private void SetCellValue(string value, OpenXmlSs.CellValues dataType = OpenXmlSs.CellValues.Error)
        {
            this.Element.CellValue = new OpenXmlSs.CellValue(value);
            if (dataType != OpenXmlSs.CellValues.Error)
            {
                this.Element.DataType = dataType;
            }
        }
        private OpenXmlSs.SharedStringItem GetSharedStringItemById(int id)
        {
            return this.Worksheet.Spreadsheet.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<OpenXmlSs.SharedStringItem>().ElementAt(id);
        }

        private delegate T ParseDelegate<T>(string s, IFormatProvider provider);

        private T GetValue<T>(Func<string, T> parse) where T : IConvertible
        {
            try
            {
                return parse(GetStringValue());
            }
            catch (FormatException)
            {
                throw;// new ApplicationServerException(MethodResult.IncorrectFormat);
            }
        }
        private T GetInvariantValue<T>(Func<string, IFormatProvider, T> parse) where T : IConvertible
        {
            try
            {
                return parse(GetStringValue(), CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                throw;// new ApplicationServerException(MethodResult.IncorrectFormat);
            }
        }
    }
}