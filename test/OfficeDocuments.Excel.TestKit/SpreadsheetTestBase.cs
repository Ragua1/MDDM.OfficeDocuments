using OfficeDocuments.Excel.Interfaces;

namespace OfficeDocuments.Excel.TestKit;

/// <summary>
/// Base for the tiers that work with a real workbook.
/// </summary>
/// <remarks>
/// Prefer <see cref="CreateInMemorySpreadsheet()"/>. A test only needs
/// <see cref="GetFilepath"/> when the path itself is part of what it is checking — opening by
/// name, a missing directory, or a file handle being released. Everything else runs faster and
/// leaks nothing when it stays in memory.
/// </remarks>
public abstract class SpreadsheetTestBase : IDisposable
{
    /// <summary>
    /// Creates a workbook backed by a fresh <see cref="MemoryStream"/>. The default choice.
    /// </summary>
    protected static ISpreadsheet CreateInMemorySpreadsheet() => Spreadsheet.CreateDocument(new MemoryStream());

    /// <summary>
    /// Creates a workbook backed by a fresh <see cref="MemoryStream"/> and hands the stream back,
    /// so the test can reopen the result or capture it with <see cref="SaveArtifact"/>.
    /// </summary>
    protected static ISpreadsheet CreateInMemorySpreadsheet(out MemoryStream stream)
    {
        stream = new MemoryStream();

        return Spreadsheet.CreateDocument(stream);
    }

    /// <summary>
    /// Writes the produced workbook out for manual inspection, but only when artifact capture is
    /// enabled — see <see cref="TestArtifacts"/>. A no-op otherwise, so it is safe to leave in
    /// place on any test whose output is worth eyeballing in Excel.
    /// </summary>
    /// <returns>The path written, or <see langword="null"/> when capture is off.</returns>
    protected string? SaveArtifact(Stream workbook, string fileName) =>
        TestArtifacts.Save(workbook, this, fileName);

    /// <summary>
    /// Creates a workbook backed by <paramref name="stream"/>, for tests that reopen it afterwards.
    /// </summary>
    protected static ISpreadsheet CreateNewSpreadsheet(Stream stream) => Spreadsheet.CreateDocument(stream);

    protected static ISpreadsheet CreateNewSpreadsheet(string filepath) => new Spreadsheet(filepath, true);

    protected static ISpreadsheet OpenExistingSpreadsheet(string filepath) => new Spreadsheet(filepath, false);

    protected static ISpreadsheet OpenExistingSpreadsheet(Stream stream) => Spreadsheet.OpenDocument(stream, true);

    /// <summary>
    /// A path inside this test class's private temp workspace. Use only when the test is about
    /// the file system; otherwise use <see cref="CreateInMemorySpreadsheet()"/>.
    /// </summary>
    protected string GetFilepath(string filename) => TempWorkspace.GetFilepath(this, filename);

    public void Dispose()
    {
        TempWorkspace.Cleanup(this);
        GC.SuppressFinalize(this);
    }
}
