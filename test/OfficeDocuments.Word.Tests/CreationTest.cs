using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDocuments.Word.Enums;

namespace OfficeDocuments.Word.Tests;

[TestClass]
public class CreationTest : TestBase
{
    public static readonly Random Rnd = new Random();

    [TestMethod]
    public void CreateEmptyDocument_EmitsNewFile()
    {
        var filename = "doc1.docx";
        CleanFilepath(filename);

        var filePath = filename;// GetFilepath(filename);
        using (var w = CreateWordProcessingDocument(filePath))
        {
            ;
        }

        Assert.IsTrue(File.Exists(filename));
    }

    [TestMethod]
    public void CreateDocumentWithContent_EmitsNewFile()
    {
        var filename = "doc2.docx";
        CleanFilepath(filename);

        var filePath = filename;// GetFilepath(filename);
        using (var w = CreateWordProcessingDocument(filePath))
        {
            var body = w.GetBody();
            body.AddParagraph()
                .AddText($"Create text on first page - {DateTime.Now:s}")
                .AddBreak(BreakType.Page);
                
            body = w.GetBody();
            body.AddParagraph()
                .AddText($"Create text on first page - {DateTime.Now:s}")
                .AddBreak(BreakType.Page)
                .AddText($"Create text on second page - {DateTime.Now:s}");
        }

        Assert.IsTrue(File.Exists(filename));
    }
}