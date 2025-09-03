using System.Linq;
using OfficeDocuments.Word.Interfaces;

namespace OfficeDocuments.Word.Tests;

public class WordprocessingTest : TestBase
{
    [Fact(Skip = "Depends on external resource file")] 
    public void ReadDocumentTest()
    {
        var path = "Resources/Rozsudek_priloha_6.docx";
        Assert.True(File.Exists(path));

        using IWordprocessing wp = new Wordprocessing(path, false);

        var body = wp.GetBody();

        var texts = body.Paragraphs.Select(x => x.GetTextElements()).Where(x => x.Any()).ToArray();
            
    Assert.True(texts.Any());

        wp.Close(false);
    }
}