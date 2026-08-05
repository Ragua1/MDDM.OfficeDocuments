using System.Runtime.CompilerServices;

namespace OfficeDocuments.Excel.TestKit;

/// <summary>
/// Per-test-class temporary directory, for the tiers that genuinely need a file on disk.
/// Tests that only need a workbook should use a <see cref="MemoryStream"/> instead.
/// </summary>
/// <remarks>
/// When <see cref="TestArtifacts"/> capture is enabled the workspace roots itself under the
/// capture directory and is not deleted, so the produced workbooks can be opened in Excel.
/// </remarks>
public static class TempWorkspace
{
    private static readonly string RootPath = Path.Combine(
        TestArtifacts.RootPath ?? Path.Combine(Path.GetTempPath(), "MDDM.OfficeDocuments.Tests"),
        "Excel");

    private static readonly ConditionalWeakTable<object, string> WorkspacePaths = new();

    /// <summary>
    /// Returns a path inside <paramref name="testClass"/>'s private workspace, creating the
    /// directory tree as needed.
    /// </summary>
    public static string GetFilepath<T>(T testClass, string filename)
    {
        ArgumentNullException.ThrowIfNull(testClass);
        ArgumentException.ThrowIfNullOrWhiteSpace(filename);

        var workspacePath = GetWorkspacePath(testClass);
        var filePath = Path.Combine(workspacePath, filename);
        var fileDirectory = Path.GetDirectoryName(filePath);

        if (fileDirectory is not null)
        {
            Directory.CreateDirectory(fileDirectory);
        }

        return filePath;
    }

    /// <summary>
    /// Best-effort removal of <paramref name="testClass"/>'s workspace. Skipped entirely when
    /// artifact capture is on, which is the whole point of turning it on.
    /// </summary>
    public static void Cleanup<T>(T testClass)
    {
        if (testClass is null || TestArtifacts.IsEnabled)
        {
            return;
        }

        if (!WorkspacePaths.TryGetValue(testClass, out var workspacePath))
        {
            return;
        }

        try
        {
            if (Directory.Exists(workspacePath))
            {
                Directory.Delete(workspacePath, true);
            }
        }
        catch (IOException)
        {
            // Keep cleanup best-effort. Temporary files may be locked by external processes.
        }
        catch (UnauthorizedAccessException)
        {
            // Keep cleanup best-effort. Temporary files may be locked by external processes.
        }

        WorkspacePaths.Remove(testClass);
    }

    private static string GetWorkspacePath<T>(T testClass)
    {
        return WorkspacePaths.GetValue(testClass!, static key =>
        {
            var typeName = TestArtifacts.SanitizePathSegment(key.GetType().Name);

            // Throwaway runs isolate every test instance; capture runs are read by a human, so
            // they get a plain per-class folder instead of a timestamped GUID.
            var workspacePath = TestArtifacts.IsEnabled
                ? Path.Combine(RootPath, typeName)
                : Path.Combine(RootPath, typeName, $"{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid():N}");

            Directory.CreateDirectory(workspacePath);

            return workspacePath;
        });
    }
}