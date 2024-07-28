using System.IO;
using OfficeDocuments.Word.Interfaces;

namespace OfficeDocuments.Word.Tests
{
    public class TestBase
    {
        protected IWordprocessing CreateWordProcessingDocument(Stream stream) => new Wordprocessing(stream, createNew: true);
        protected IWordprocessing CreateWordProcessingDocument(string filepath) => new Wordprocessing(filepath, true);

        protected IWordprocessing OpenWordProcessingDocument(string filepath) => new Wordprocessing(filepath, false);
        protected IWordprocessing OpenWordProcessingDocument(Stream stream) => new Wordprocessing(stream, false);


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
