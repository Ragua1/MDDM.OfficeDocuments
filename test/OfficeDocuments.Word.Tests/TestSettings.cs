namespace OfficeDocuments.Word.Tests;

public static class TestSettings
{
    internal static string GetFilepath<T>(T testClass, string filename)
    {
        ArgumentNullException.ThrowIfNull(testClass);

        var path = Path.Combine(Path.GetTempPath(), testClass.GetType().Name);

        if (!Directory.Exists(path))
        {
            Directory.CreateDirectory(path);
        }

        return Path.Combine(path, filename);
    }
}