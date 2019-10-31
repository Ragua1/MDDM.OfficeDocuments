using System.IO;

namespace OfficeDocumentsApi.Excel.Test
{
    public static class TestSettings
    {
        internal static string GetFilepath<T>(T testClass, string filename)
        {
            var path = Path.Combine(Path.GetTempPath(), testClass.GetType().Name);

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            return Path.Combine(path, filename);
        }
    }
}