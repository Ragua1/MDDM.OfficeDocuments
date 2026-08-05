using OfficeDocuments.Excel.Extensions;

namespace OfficeDocuments.Excel.UnitTests;

public class CellExtensionTests
{
    [Theory]
    [InlineData("A", 1)]
    [InlineData("Z", 26)]
    [InlineData("AA", 27)]
    [InlineData("az", 52)]
    public void GetExcelColumnIndex_ValidColumnName_ReturnsExpectedIndex(string columnName, uint expectedIndex)
    {
        var columnIndex = columnName.GetExcelColumnIndex();

        Assert.Equal(expectedIndex, columnIndex);
    }

    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(52, "AZ")]
    public void GetExcelColumnName_ValidColumnIndex_ReturnsExpectedName(uint columnIndex, string expectedName)
    {
        var columnName = columnIndex.GetExcelColumnName();

        Assert.Equal(expectedName, columnName);
    }

    [Theory]
    [InlineData(0)]
    public void GetExcelColumnName_InvalidColumnIndex_ThrowsArgumentException(uint columnIndex)
    {
        Assert.Throws<ArgumentException>(() => columnIndex.GetExcelColumnName());
    }

    [Theory]
    [InlineData(1, 1, "A1")]
    [InlineData(28, 12, "AB12")]
    public void GetExcelCellReference_ValidCoordinates_ReturnsExpectedReference(uint columnIndex, uint rowIndex, string expectedReference)
    {
        var reference = CellExtension.GetExcelCellReference(columnIndex, rowIndex);

        Assert.Equal(expectedReference, reference);
    }

    [Theory]
    [InlineData("A1", 1, 1)]
    [InlineData("ab12", 12, 28)]
    [InlineData("  C4  ", 4, 3)]
    public void GetExcelCellIndex_ValidReference_ReturnsCoordinates(string reference, uint expectedRowIndex, uint expectedColumnIndex)
    {
        var (rowIndex, columnIndex) = reference.GetExcelCellIndex();

        Assert.Equal(expectedRowIndex, rowIndex);
        Assert.Equal(expectedColumnIndex, columnIndex);
    }

    [Theory]
    [InlineData("")]
    [InlineData("A0")]
    [InlineData("1A")]
    [InlineData("A 1")]
    [InlineData("A:1")]
    [InlineData("A_1")]
    public void GetExcelCellIndex_InvalidReference_ThrowsArgumentException(string reference)
    {
        Assert.Throws<ArgumentException>(() => reference.GetExcelCellIndex());
    }

    [Theory]
    [InlineData("A1", 1, 1, 1, 1)]
    [InlineData("A1:B3", 1, 1, 2, 3)]
    [InlineData("  c4:d5  ", 3, 4, 4, 5)]
    public void TryGetExcelRange_ValidReference_ReturnsCoordinates(string reference, uint expectedFromColumn, uint expectedFromRow, uint expectedToColumn, uint expectedToRow)
    {
        var result = reference.TryGetExcelRange(out var coordinates);

        Assert.True(result);
        Assert.Equal((expectedFromColumn, expectedFromRow, expectedToColumn, expectedToRow), coordinates);
    }

    [Theory]
    [InlineData("")]
    [InlineData("A0")]
    [InlineData("A1:B0")]
    [InlineData("A1:B2:C3")]
    [InlineData("A1:")]
    [InlineData(":B2")]
    public void TryGetExcelRange_InvalidReference_ReturnsFalse(string reference)
    {
        var result = reference.TryGetExcelRange(out var coordinates);

        Assert.False(result);
        Assert.Equal(default, coordinates);
    }
}