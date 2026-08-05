using OfficeDocuments.Excel.TestKit;

namespace OfficeDocuments.Excel.IntegrationTests;

/// <summary>
/// Worksheet naming through the public API (EXCEL-011 phase 6, blind spot B-2). The rule table
/// itself is covered in the unit tier; what matters here is that every entry point applies it —
/// creating, renaming, and copying a sheet all end up writing the same attribute.
/// </summary>
public class WorksheetNamingTests : SpreadsheetTestBase
{
    [Theory]
    [InlineData("Report/2026")]
    [InlineData("A:B")]
    [InlineData("what?")]
    [InlineData("[Book]")]
    [InlineData("a*b")]
    [InlineData("back\\slash")]
    public void AddWorksheet_RejectsNamesExcelForbids(string name)
    {
        using var workbook = CreateInMemorySpreadsheet();

        Assert.Throws<ArgumentException>(() => workbook.AddWorksheet(name));
    }

    [Fact]
    public void AddWorksheet_RejectsNamesLongerThanThirtyOneCharacters()
    {
        using var workbook = CreateInMemorySpreadsheet();

        workbook.AddWorksheet(new string('a', 31));
        Assert.Throws<ArgumentException>(() => workbook.AddWorksheet(new string('b', 32)));
    }

    [Fact]
    public void RenameWorksheet_AppliesTheSameRules()
    {
        using var workbook = CreateInMemorySpreadsheet();
        workbook.AddWorksheet("Data");

        Assert.Throws<ArgumentException>(() => workbook.RenameWorksheet("Data", "Data/2026"));
        Assert.Throws<ArgumentException>(() => workbook.RenameWorksheet("Data", new string('x', 40)));
    }

    /// <summary>
    /// Excel matches sheet names without regard to case, so <c>DATA</c> and <c>Data</c> cannot
    /// coexist. This already held; it is pinned here because the validation added around it is
    /// exactly the kind of change that could have replaced the comparison.
    /// </summary>
    [Fact]
    public void AddWorksheet_TreatsNamesAsCaseInsensitiveForUniqueness()
    {
        using var workbook = CreateInMemorySpreadsheet();
        workbook.AddWorksheet("Data");

        Assert.Throws<ArgumentException>(() => workbook.AddWorksheet("DATA"));
    }

    /// <summary>
    /// An apostrophe is fine inside the name and not at the ends, where it would collide with the
    /// quoting in a <c>'Sheet Name'!A1</c> reference.
    /// </summary>
    [Fact]
    public void AddWorksheet_AllowsAnApostropheOnlyInsideTheName()
    {
        using var workbook = CreateInMemorySpreadsheet();

        workbook.AddWorksheet("Bob's data");
        Assert.Throws<ArgumentException>(() => workbook.AddWorksheet("'quoted'"));
    }

    /// <summary>
    /// Markup characters are legal in a sheet name; the SDK escapes them into
    /// <c>workbook.xml</c>. This is the case that corrupted files in other libraries, which built
    /// that attribute by string concatenation.
    /// </summary>
    [Fact]
    public void AddWorksheet_RoundTripsMarkupCharactersInTheName()
    {
        const string name = "R&D <2026>";

        using var workbook = CreateInMemorySpreadsheet(out var stream);
        workbook.AddWorksheet(name).AddCell(1u, 1u, "value");
        workbook.Close();

        stream.Position = 0;
        using var reopened = OpenExistingSpreadsheet(stream);
        Assert.Contains(name, reopened.GetWorksheetsName());
        Assert.Equal("value", reopened.GetWorksheet(name)!.GetCell(1u, 1u)!.GetStringValue());
    }

    /// <summary>The default name the library generates when the caller supplies none must be legal too.</summary>
    [Fact]
    public void AddWorksheet_GeneratesALegalDefaultName()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var worksheet = workbook.AddWorksheet();

        Assert.InRange(worksheet.Name.Length, 1, 31);
        Assert.DoesNotContain(worksheet.Name, name => ":\\/?*[]".Contains(name));
    }
}
