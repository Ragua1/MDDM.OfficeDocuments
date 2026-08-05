namespace OfficeDocuments.Excel.TestKit;

/// <summary>
/// Opt-in capture of the workbooks a test produces, so they can be opened in Excel and inspected
/// by hand.
/// </summary>
/// <remarks>
/// Off by default: tests run in memory and leave nothing behind. Set the
/// <c>OFFICEDOCS_TEST_OUTPUT</c> environment variable to turn capture on:
/// <list type="bullet">
/// <item><description><c>1</c> / <c>true</c> — write under <c>%TEMP%/MDDM.OfficeDocuments.Tests/Output</c>.</description></item>
/// <item><description>any other value — treat it as the target directory.</description></item>
/// </list>
/// When capture is on, <see cref="TempWorkspace"/> also roots itself under the same directory and
/// stops deleting itself, so every test that writes a real file leaves it there automatically.
/// </remarks>
public static class TestArtifacts
{
    /// <summary>
    /// Environment variable that enables capture and optionally names the target directory.
    /// </summary>
    public const string EnvironmentVariable = "OFFICEDOCS_TEST_OUTPUT";

    /// <summary>
    /// Directory that captured workbooks are written to, or <see langword="null"/> when capture
    /// is off.
    /// </summary>
    public static string? RootPath { get; } = ResolveRootPath();

    /// <summary>
    /// Whether capture is enabled for this test run.
    /// </summary>
    public static bool IsEnabled => RootPath is not null;

    /// <summary>
    /// Writes <paramref name="workbook"/> to the capture directory when capture is on, and does
    /// nothing otherwise. The stream position is restored before returning.
    /// </summary>
    /// <returns>The path written, or <see langword="null"/> when capture is off.</returns>
    public static string? Save(Stream workbook, object testClass, string fileName)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(testClass);
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);

        if (RootPath is null)
        {
            return null;
        }

        var targetPath = ResolveTargetPath(testClass, fileName);
        var originalPosition = workbook.Position;
        workbook.Position = 0;

        try
        {
            using var file = File.Create(targetPath);
            workbook.CopyTo(file);
        }
        finally
        {
            workbook.Position = originalPosition;
        }

        return targetPath;
    }

    /// <summary>
    /// Copies an already-written workbook into the capture directory when capture is on.
    /// </summary>
    /// <returns>The path written, or <see langword="null"/> when capture is off.</returns>
    public static string? Save(string sourceFilePath, object testClass, string? fileName = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceFilePath);
        ArgumentNullException.ThrowIfNull(testClass);

        if (RootPath is null)
        {
            return null;
        }

        var targetPath = ResolveTargetPath(testClass, fileName ?? Path.GetFileName(sourceFilePath));
        File.Copy(sourceFilePath, targetPath, overwrite: true);

        return targetPath;
    }

    internal static string SanitizePathSegment(string value)
    {
        var invalidChars = Path.GetInvalidFileNameChars();

        return new string(value.Select(character => invalidChars.Contains(character) ? '_' : character).ToArray());
    }

    private static string ResolveTargetPath(object testClass, string fileName)
    {
        var directory = Path.Combine(RootPath!, "Excel", SanitizePathSegment(testClass.GetType().Name));
        Directory.CreateDirectory(directory);

        return Path.Combine(directory, SanitizePathSegment(fileName));
    }

    private static string? ResolveRootPath()
    {
        var value = Environment.GetEnvironmentVariable(EnvironmentVariable);
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        value = value.Trim();

        if (string.Equals(value, "0", StringComparison.Ordinal)
            || string.Equals(value, "false", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (string.Equals(value, "1", StringComparison.Ordinal)
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase))
        {
            return Path.Combine(Path.GetTempPath(), "MDDM.OfficeDocuments.Tests", "Output");
        }

        return Path.GetFullPath(value);
    }
}
