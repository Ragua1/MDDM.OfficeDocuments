using System.IO;

namespace OfficeDocumentsApi.Test.TestBases
{
    public class WordprocessingTestBase
    {
        protected OfficeDocumentsApi.Wordprocessing CreateTestee(Stream stream) => new OfficeDocumentsApi.Wordprocessing(stream, true);
        protected OfficeDocumentsApi.Wordprocessing CreateTestee(string filepath) => new OfficeDocumentsApi.Wordprocessing(filepath, true);

        protected OfficeDocumentsApi.Wordprocessing CreateOpenTestee(string filepath) => new OfficeDocumentsApi.Wordprocessing(filepath, false);
        protected OfficeDocumentsApi.Wordprocessing CreateOpenTestee(Stream stream) => new OfficeDocumentsApi.Wordprocessing(stream, false);


        protected string GetFilepath(string filename) => TestSettings.GetFilepath(this, filename);
        protected void CleanFilepath(string filename)
        {
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }
        }
    }
}
