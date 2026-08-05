using OfficeDocuments.Excel.TestKit;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Styles;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.IntegrationTests;

public class WriterTest : SpreadsheetTestBase
{
    [Fact]
    public void CreateNewFile()
    {
        var filePath = GetFilepath("doc1.xlsx");

        using var _ = CreateNewSpreadsheet(filePath);

        Assert.True(File.Exists(filePath), $"File not exist on filePath: '{filePath}'");
    }

    [Fact]
    public void CreateStyleWithDefaultSettings()
    {
        var filePath = GetFilepath("doc2.xlsx");

        using (var writer = CreateNewSpreadsheet(filePath))
        {
            var style = writer.CreateStyle();

            Assert.Equal(0, style.FontId);
            Assert.Equal(0, style.FillId);
            Assert.Equal(0, style.BorderId);
            Assert.Equal(0, style.NumberFormatId);

            Assert.Equal((uint)0, style.StyleIndex);
        }
    }

    [Fact]
    public void CreateStyle()
    {
        var filePath = GetFilepath("doc3.xlsx");

        using (var writer = CreateNewSpreadsheet(filePath))
        {
            var style = writer.CreateStyle(
                new Font { FontSize = 12, Color = Color.Aqua, FontName = FontNameValues.Tahoma },
                new Fill(Color.Brown),
                new Border { Left = BorderStyleValues.Double, Right = BorderStyleValues.Double }
            );

            Assert.True(style.FontId > 0);
            Assert.True(style.FillId > 0);
            Assert.True(style.BorderId > 0);
            Assert.Equal(0, style.NumberFormatId);

            Assert.True(style.StyleIndex > 0);
        }
    }

    [Fact]
    public void NotCreateFileOnNonExistDirectory()
    {
        var filePath = Path.Combine(GetFilepath("missing"), "doc4.xlsx");

        Assert.Throws<DirectoryNotFoundException>(() =>
        {
            using var _ = CreateNewSpreadsheet(filePath);
            Assert.Fail("Should not reach this point");
        });

        Assert.False(File.Exists(filePath), $"File should not exist on filePath: '{filePath}'");
    }

    [Fact]
    public void GetExistedWorksheet()
    {
        var filePath = GetFilepath("doc5.xlsx");

        using (var writer = CreateNewSpreadsheet(filePath))
        {
            var sheetName = "Test1";
            var sheet1 = writer.AddWorksheet(sheetName);

            var sheet2 = writer.GetWorksheet(sheetName);

            Assert.NotNull(sheet2);
            Assert.Same(sheet1, sheet2);
        }
    }

    [Fact]
    public void OpenExistedDocument()
    {
        var filePath = GetFilepath("doc6.xlsx");

        using (var writer = CreateNewSpreadsheet(filePath))
        {
            writer.AddWorksheet("Test1");
        }

        using (var writer = OpenExistingSpreadsheet(filePath))
        {
            Assert.True(writer.GetWorksheetsName().Any());
        }
    }

    [Fact]
    public void OpenTryFindNonExistedWorksheet()
    {
        var filePath = GetFilepath("doc7.xlsx");

        using (var writer = CreateNewSpreadsheet(filePath))
        {
            writer.AddWorksheet("Test1");
        }

        using (var writer = OpenExistingSpreadsheet(filePath))
        {
            Assert.True(writer.GetWorksheetsName().Any());
            Assert.Null(writer.GetWorksheet("Test"));
        }
    }

    [Fact]
    public void CreateDocumentToStream()
    {
        var stream = new MemoryStream();

        using (var writer = CreateNewSpreadsheet(stream))
        {
            writer.AddWorksheet("Test1");
        }
    }

    [Fact]
    public void OpenDocumentFromStream()
    {
        var stream = new MemoryStream();

        using (var writer = CreateNewSpreadsheet(stream))
        {
            writer.AddWorksheet("Test1");
        }

        using (var writer = OpenExistingSpreadsheet(stream))
        {
            Assert.True(writer.GetWorksheetsName().Any());
        }
    }
}