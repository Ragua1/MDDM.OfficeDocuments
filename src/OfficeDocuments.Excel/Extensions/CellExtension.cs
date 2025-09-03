namespace OfficeDocuments.Excel.Extensions;

public static class CellExtension
{
    public static uint GetExcelColumnIndex(this string columnName)
    {
        return (uint)columnName
            .ToUpper()
            .Aggregate(0, (column, letter) => 26 * column + letter - 'A' + 1);
    }
    public static (uint rowIndex, uint columnIndex) GetExcelCellIndex(this string cellReference)
    {
        return (
            uint.Parse(new string(cellReference.Where(char.IsDigit).ToArray())),
            new string(cellReference.Where(char.IsLetter).ToArray()).GetExcelColumnIndex()
            );
    }
}
