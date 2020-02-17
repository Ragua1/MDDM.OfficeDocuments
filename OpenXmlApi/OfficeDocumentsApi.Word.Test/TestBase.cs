using System.IO;
using OfficeDocumentsApi.Word;
using OfficeDocumentsApi.Word.Interfaces;

namespace OfficeDocumentsApi.Word.Test
{
    public class TestBase
    {
        protected IWordprocessing CreateTestee(Stream stream) => new Word.Wordprocessing(stream, true);
        protected IWordprocessing CreateTestee(string filepath) => new Word.Wordprocessing(filepath, true);

        protected IWordprocessing CreateOpenTestee(string filepath) => new Word.Wordprocessing(filepath, false);
        protected IWordprocessing CreateOpenTestee(Stream stream) => new Word.Wordprocessing(stream, false);


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
