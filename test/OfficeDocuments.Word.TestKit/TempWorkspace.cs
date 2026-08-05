using System.Runtime.CompilerServices;

namespace OfficeDocuments.Word.TestKit;

/// <summary>
/// Per-test-class temporary directory, for the tests that genuinely need a file on disk.
/// Tests that only need a document should use a <see cref="MemoryStream"/> instead.
/// </summary>
public static class TempWorkspace
{
    private static readonly string RootPath = Path.Combine(Path.GetTempPath(), "MDDM.OfficeDocuments.Tests", "Word");
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
    /// Best-effort removal of <paramref name="testClass"/>'s workspace.
    /// </summary>
    public static void Cleanup<T>(T testClass)
    {
        if (testClass is null)
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
            var typeName = SanitizePathSegment(key.GetType().Name);
            var uniqueFolder = $"{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid():N}";
            var workspacePath = Path.Combine(RootPath, typeName, uniqueFolder);
            Directory.CreateDirectory(workspacePath);
            return workspacePath;
        });
    }

    private static string SanitizePathSegment(string value)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitizedChars = value
            .Select(ch => invalidChars.Contains(ch) ? '_' : ch)
            .ToArray();

        return new string(sanitizedChars);
    }
}
