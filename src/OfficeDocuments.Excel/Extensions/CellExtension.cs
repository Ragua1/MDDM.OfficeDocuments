using System.Globalization;

namespace OfficeDocuments.Excel.Extensions;

public static class CellExtension
{
    extension(string value)
    {
        public uint GetExcelColumnIndex()
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("Column name cannot be null or empty.", nameof(value));
            }

            if (!TryParseColumnName(value.AsSpan(), out var columnIndex))
            {
                throw new ArgumentException($"Invalid column name '{value}'", nameof(value));
            }

            return columnIndex;
        }

        public (uint rowIndex, uint columnIndex) GetExcelCellIndex()
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("Cell reference cannot be null or empty.", nameof(value));
            }

            if (!TryParseCellReference(value.AsSpan(), out var rowIndex, out var columnIndex))
            {
                throw new ArgumentException($"Invalid cell reference '{value}'", nameof(value));
            }

            return (rowIndex, columnIndex);
        }

        public bool TryGetExcelRange(out (uint fromColumn, uint fromRow, uint toColumn, uint toRow) coordinates)
        {
            return TryParseRange(value.AsSpan(), out coordinates);
        }
    }

    extension(uint columnIndex)
    {
        public string GetExcelColumnName()
        {
            if (columnIndex < 1)
            {
                throw new ArgumentException($"Invalid argument column index '{columnIndex}'", nameof(columnIndex));
            }

            Span<char> buffer = stackalloc char[8];
            var position = buffer.Length;
            var dividend = columnIndex;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                buffer[--position] = (char)('A' + modulo);
                dividend = (dividend - 1) / 26;
            }

            return new string(buffer[position..]);
        }
    }

    public static string GetExcelCellReference(uint columnIndex, uint rowIndex)
    {
        if (columnIndex < 1)
        {
            throw new ArgumentException($"Invalid argument column index '{columnIndex}'", nameof(columnIndex));
        }

        if (rowIndex < 1)
        {
            throw new ArgumentException($"Invalid argument row index '{rowIndex}'", nameof(rowIndex));
        }

        return string.Concat(columnIndex.GetExcelColumnName(), rowIndex.ToString(CultureInfo.InvariantCulture));
    }

    internal static bool TryParseCellReference(ReadOnlySpan<char> cellReference, out uint rowIndex, out uint columnIndex)
    {
        rowIndex = 0;
        columnIndex = 0;

        cellReference = cellReference.Trim();
        if (cellReference.IsEmpty)
        {
            return false;
        }

        var splitIndex = 0;
        while (splitIndex < cellReference.Length && TryNormalizeLetter(cellReference[splitIndex], out _))
        {
            splitIndex++;
        }

        if (splitIndex == 0 || splitIndex == cellReference.Length)
        {
            return false;
        }

        if (!TryParseColumnName(cellReference[..splitIndex], out columnIndex))
        {
            return false;
        }

        ulong parsedRow = 0;
        foreach (var character in cellReference[splitIndex..])
        {
            if (!char.IsAsciiDigit(character))
            {
                rowIndex = 0;
                columnIndex = 0;
                return false;
            }

            parsedRow = (parsedRow * 10) + (ulong)(character - '0');
            if (parsedRow > uint.MaxValue)
            {
                rowIndex = 0;
                columnIndex = 0;
                return false;
            }
        }

        if (parsedRow == 0)
        {
            rowIndex = 0;
            columnIndex = 0;
            return false;
        }

        rowIndex = (uint)parsedRow;
        return true;
    }

    private static bool TryParseRange(ReadOnlySpan<char> rangeReference, out (uint fromColumn, uint fromRow, uint toColumn, uint toRow) coordinates)
    {
        coordinates = default;

        rangeReference = rangeReference.Trim();
        if (rangeReference.IsEmpty)
        {
            return false;
        }

        var separatorIndex = rangeReference.IndexOf(':');
        if (separatorIndex >= 0 && separatorIndex != rangeReference.LastIndexOf(':'))
        {
            return false;
        }

        if (separatorIndex < 0)
        {
            if (!TryParseCellReference(rangeReference, out var rowIndex, out var columnIndex))
            {
                return false;
            }

            coordinates = (columnIndex, rowIndex, columnIndex, rowIndex);
            return true;
        }

        var fromReference = rangeReference[..separatorIndex].Trim();
        var toReference = rangeReference[(separatorIndex + 1)..].Trim();
        if (!TryParseCellReference(fromReference, out var fromRow, out var fromColumn)
            || !TryParseCellReference(toReference, out var toRow, out var toColumn))
        {
            return false;
        }

        coordinates = (fromColumn, fromRow, toColumn, toRow);
        return true;
    }

    private static bool TryParseColumnName(ReadOnlySpan<char> columnName, out uint columnIndex)
    {
        columnIndex = 0;
        if (columnName.IsEmpty)
        {
            return false;
        }

        ulong parsedColumnIndex = 0;
        foreach (var character in columnName)
        {
            if (!TryNormalizeLetter(character, out var normalizedLetter))
            {
                columnIndex = 0;
                return false;
            }

            parsedColumnIndex = (parsedColumnIndex * 26) + (uint)(normalizedLetter - 'A' + 1);
            if (parsedColumnIndex > uint.MaxValue)
            {
                columnIndex = 0;
                return false;
            }
        }

        columnIndex = (uint)parsedColumnIndex;
        return true;
    }

    private static bool TryNormalizeLetter(char character, out char normalizedLetter)
    {
        if (character is >= 'A' and <= 'Z')
        {
            normalizedLetter = character;
            return true;
        }

        if (character is >= 'a' and <= 'z')
        {
            normalizedLetter = (char)(character - ('a' - 'A'));
            return true;
        }

        normalizedLetter = default;
        return false;
    }
}
