using BenchmarkDotNet.Attributes;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;

namespace OfficeDocuments.Excel.Benchmarks;

/// <summary>
/// <c>Range.SortByColumn</c>. The sort itself is a normal comparison sort, but it snapshots the
/// range by deep-cloning every cell element, then writes every cell back through
/// <c>AddCellOnIndex</c>. The clone-and-replay dominates; the comparisons do not.
/// <para>
/// Sorting mutates the sheet, so it cannot be measured against a shared fixture. Both benchmarks
/// build their own data and the baseline exists so the build cost can be subtracted.
/// </para>
/// </summary>
[Config(typeof(ScalingConfig))]
public class RangeSortBenchmarks
{
    private const uint Columns = 5;

    [Params(500, 1_000, 2_000)]
    public int Rows;

    [Benchmark(Description = "build N rows", Baseline = true)]
    public int Build()
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        Populate(spreadsheet, Rows);
        return Rows;
    }

    [Benchmark(Description = "build N rows + SortByColumn")]
    public int BuildAndSort()
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        var worksheet = Populate(spreadsheet, Rows);

        worksheet.GetRange(1, 1, Columns, (uint)Rows)
            .SortByColumn(1, SortDirection.Ascending);

        return Rows;
    }

    /// <summary>
    /// Descending values, so the sort has to move essentially every row — an already-sorted
    /// input would let the comparison sort finish early and understate the replay cost.
    /// </summary>
    private static IWorksheet Populate(ISpreadsheet spreadsheet, int rows)
    {
        var worksheet = spreadsheet.AddWorksheet("Sheet");

        for (var rowIndex = 1u; rowIndex <= rows; rowIndex++)
        {
            worksheet.AddCell(1u, rowIndex, rows - (int)rowIndex);
            for (var column = 2u; column <= Columns; column++)
            {
                worksheet.AddCell(column, rowIndex, $"r{rowIndex}c{column}");
            }
        }

        return worksheet;
    }
}
