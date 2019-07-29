using System;
using OpenXmlApi.Interfaces;

namespace OpenXmlApi.DataClasses
{
    public class Cell : ICell
    {
        public DocumentFormat.OpenXml.Spreadsheet.Cell Element => throw new NotImplementedException();

        public string CellReference => throw new NotImplementedException();

        public uint RowIndex => throw new NotImplementedException();

        public uint ColumnIndex => throw new NotImplementedException();

        public string Value { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public IWorksheet Worksheet => throw new NotImplementedException();

        public IStyle Style => throw new NotImplementedException();

        public IStyle AddStyle(params IStyle[] styles)
        {
            throw new NotImplementedException();
        }

        public bool GetBoolValue()
        {
            throw new NotImplementedException();
        }

        public DateTime GetDateValue(string format = null)
        {
            throw new NotImplementedException();
        }

        public decimal GetDecimalValue()
        {
            throw new NotImplementedException();
        }

        public double GetDoubleValue()
        {
            throw new NotImplementedException();
        }

        public string GetFormula()
        {
            throw new NotImplementedException();
        }

        public int GetIntValue()
        {
            throw new NotImplementedException();
        }

        public long GetLongValue()
        {
            throw new NotImplementedException();
        }

        public string GetStringValue()
        {
            throw new NotImplementedException();
        }

        public bool HasFormula()
        {
            throw new NotImplementedException();
        }

        public bool HasValue()
        {
            throw new NotImplementedException();
        }

        public void SetFormula(string formula)
        {
            throw new NotImplementedException();
        }

        public void SetValue(object value)
        {
            throw new NotImplementedException();
        }

        public void SetValue(bool value)
        {
            throw new NotImplementedException();
        }

        public void SetValue(int value)
        {
            throw new NotImplementedException();
        }

        public void SetValue(long value)
        {
            throw new NotImplementedException();
        }

        public void SetValue(double value)
        {
            throw new NotImplementedException();
        }

        public void SetValue(decimal value)
        {
            throw new NotImplementedException();
        }

        public void SetValue(DateTime value)
        {
            throw new NotImplementedException();
        }

        public void SetValue(string value)
        {
            throw new NotImplementedException();
        }

        public bool TryGetValue(out bool value)
        {
            throw new NotImplementedException();
        }

        public bool TryGetValue(out int value)
        {
            throw new NotImplementedException();
        }

        public bool TryGetValue(out long value)
        {
            throw new NotImplementedException();
        }

        public bool TryGetValue(out double value)
        {
            throw new NotImplementedException();
        }

        public bool TryGetValue(out decimal value)
        {
            throw new NotImplementedException();
        }

        public bool TryGetValue(out string value)
        {
            throw new NotImplementedException();
        }

        public bool TryGetValue(out DateTime value, string format = null)
        {
            throw new NotImplementedException();
        }
    }
}
