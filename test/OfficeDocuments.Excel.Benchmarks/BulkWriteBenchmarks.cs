using BenchmarkDotNet.Attributes;

namespace OfficeDocuments.Excel.Benchmarks;

/// <summary>
/// The whole-document path a report generator actually exercises: write N rows, close the
/// package, reopen it and read back. No known hot spot lives here — the point is to have a
/// headline number for "how large a workbook is this library practical for", and to notice if a
/// change elsewhere makes the ordinary case slower.
/// </summary>
[Config(typeof(ScalingConfig))]
public class BulkWriteBenchmarks
{
    [Params(2_000, 5_000, 10_000)]
    public int Rows;

    private Record[] _records = [];

    [GlobalSetup]
    public void Setup()
    {
        _records = Enumerable.Range(1, Rows)
            .Select(i => new Record(i, $"Item {i}", i * 1.5m, new DateTime(2026, 1, 1).AddDays(i % 365)))
            .ToArray();
    }

    /// <summary>Reflection-driven bulk insert, the one-call import path.</summary>
    [Benchmark(Description = "AddRows<T>")]
    public int AddRowsTyped()
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        var worksheet = spreadsheet.AddWorksheet("Sheet");

        worksheet.AddRows(_records, includeHeader: true);

        return Rows;
    }

    /// <summary>
    /// Cell-by-cell writes of the same data. Compared with the benchmark above this shows what
    /// the reflection in <c>AddRows&lt;T&gt;</c> costs.
    /// </summary>
    [Benchmark(Description = "AddCell per field", Baseline = true)]
    public int AddCellsManually()
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        var worksheet = spreadsheet.AddWorksheet("Sheet");

        for (var i = 0; i < _records.Length; i++)
        {
            var rowIndex = (uint)i + 1;
            var record = _records[i];
            worksheet.AddCell(1u, rowIndex, record.Id);
            worksheet.AddCell(2u, rowIndex, record.Name);
            worksheet.AddCell(3u, rowIndex, record.Amount);
            worksheet.AddCell(4u, rowIndex, record.Date);
        }

        return Rows;
    }

    /// <summary>
    /// Write, close, reopen, read every value back. This is the number to quote when someone
    /// asks how long a report of this size takes end to end.
    /// </summary>
    [Benchmark(Description = "write + close + reopen + read")]
    public int RoundTrip()
    {
        using var stream = new MemoryStream();

        using (var spreadsheet = Spreadsheet.CreateDocument(stream))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet");
            worksheet.AddRows(_records, includeHeader: true);
            spreadsheet.Close();
        }

        stream.Position = 0;

        using var reopened = Spreadsheet.OpenDocument(stream, isEditable: false);
        var sheet = reopened.GetWorksheet("Sheet")!;

        var read = 0;
        for (var rowIndex = 2u; rowIndex <= Rows + 1; rowIndex++)
        {
            if (!string.IsNullOrEmpty(sheet.GetCell(2u, rowIndex)?.GetStringValue()))
            {
                read++;
            }
        }

        return read;
    }

    public sealed record Record(int Id, string Name, decimal Amount, DateTime Date);
}
