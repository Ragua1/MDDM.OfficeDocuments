using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.TestKit;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// End-to-end scenarios that exercise the authoring surface the way a consumer would.
/// </summary>
public class WordprocessingTest : WordTestBase
{
    /// <summary>
    /// The kind of document the module exists to produce: a title, headings, formatted body text, and
    /// a page break — written once, validated against the schema, and read back.
    /// </summary>
    [Fact]
    public void AuthorReport_ProducesSchemaValidDocumentWithExpectedContent()
    {
        var bodyText = new TextFormat { FontName = "Calibri", FontSize = 11 };
        var justified = new ParagraphFormat { Alignment = ParagraphAlignment.Justify, SpacingAfter = 6 };

        using var stream = WriteAndValidate(body =>
        {
            body.AddParagraph("Quarterly report", new ParagraphFormat { StyleId = WordStyleIds.Title });
            body.AddHeading("Summary", 1);
            body.AddParagraph("Revenue grew steadily.", justified, bodyText);

            body.AddParagraph()
                .ApplyFormat(justified)
                .AddText("Total: ", bodyText)
                .AddText("1 240 000 CZK", bodyText with { Bold = true })
                .AddText(" (unaudited)", bodyText with { Italic = true });

            body.AddHeading("Detail", 2);
            body.AddParagraph().AddBreak(BreakType.Page);
            body.AddParagraph("Continued on the next page.", justified, bodyText);
        });

        using var document = OpenDocument(stream, isEditable: false);
        var paragraphs = document.GetBody().Paragraphs;

        Assert.Equal(7, paragraphs.Count);
        Assert.Equal("Quarterly report", paragraphs[0].GetTexts());
        Assert.Equal(WordStyleIds.Title, paragraphs[0].Format.StyleId);
        Assert.Equal(WordStyleIds.Heading1, paragraphs[1].Format.StyleId);
        Assert.Equal("Total: 1 240 000 CZK (unaudited)", paragraphs[3].GetTexts());
        Assert.True(paragraphs[3].Runs[1].Format.Bold);
        Assert.True(paragraphs[3].Runs[2].Format.Italic);
    }

    /// <summary>
    /// Every feature at once, in the shape a branded business document actually takes. Individual tests
    /// prove each part; this one proves they compose into one schema-valid document, which is where
    /// part ownership, relationship scope, and child ordering all have to hold together.
    /// </summary>
    [Fact]
    public void AuthorBrandedDocument_CombinesEveryFeatureAndStaysValid()
    {
        var bodyText = new TextFormat { FontName = "Calibri", FontSize = 11 };
        var filePath = GetFilepath("branded.docx");

        using (var document = CreateDocument(filePath))
        {
            document
                .ApplyPageSetup(new PageSetup { PaperSize = PaperSize.A4, Orientation = PageOrientation.Portrait }
                    .WithUniformMargins(56))
                .SetMetadata(new DocumentMetadata
                {
                    Title = "Service report",
                    Author = "MDDM",
                    Subject = "Monthly summary",
                });

            var header = document.AddHeader();
            using (var logo = new MemoryStream(TestImages.PngWithSize(120, 40)))
            {
                header.AddParagraph(new ParagraphFormat { Alignment = ParagraphAlignment.Right })
                    .AddImage(logo, ImageSize.FromWidth(60));
            }

            document.AddFooter()
                .AddParagraph("Confidential", new ParagraphFormat { Alignment = ParagraphAlignment.Center },
                    bodyText with { FontSize = 8, Color = "808080" });

            var body = document.GetBody();
            body.AddParagraph("Service report", new ParagraphFormat { StyleId = WordStyleIds.Title });
            body.AddHeading("Summary", 1);
            body.AddParagraph("Two incidents were resolved this month.", null, bodyText);

            body.AddListItem("Incident 1 — resolved", ListStyle.Number);
            body.AddListItem("Incident 2 — resolved", ListStyle.Number);

            body.AddHeading("Detail", 1);

            var table = body.AddTable([
                ["Incident", "Opened", "Closed"],
                ["INC-1", "2026-07-02", "2026-07-03"],
                ["INC-2", "2026-07-19", "2026-07-21"],
            ], new TableFormat { WidthPercent = 100, Borders = TableBorders.All, CellPadding = 3 });

            table.Rows[0].RepeatAsHeader();
            table.Rows[0].Cells[0].ApplyFormat(new TableCellFormat { BackgroundColor = "D9E2F3" });

            body.AddParagraph(new ParagraphFormat { SpacingBefore = 12, PageBreakBefore = true })
                .AddText("Full history: ", bodyText)
                .AddHyperlink("the incident portal", "https://example.com/incidents")
                .AddText(" (login required)", bodyText with { Italic = true });

            body.AddParagraph()
                .AddText("Uptime was 99.95", bodyText)
                .AddText("%", bodyText with { VerticalPosition = TextVerticalPosition.Superscript });
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenDocument(filePath, isEditable: false);

        Assert.Equal("Service report", reopened.Metadata.Title);
        Assert.Equal(PaperSize.A4, reopened.PageSetup.PaperSize);
        Assert.Equal(56, reopened.PageSetup.MarginLeft);

        var reopenedBody = reopened.GetBody();
        Assert.Single(reopenedBody.Tables);
        Assert.Equal(3, reopenedBody.Tables[0].Rows.Count);
        Assert.True(reopenedBody.Tables[0].Rows[0].IsHeader);
        Assert.Contains("the incident portal", reopenedBody.GetAllTexts(), StringComparison.Ordinal);
        Assert.Contains("INC-1\t2026-07-02", reopenedBody.GetAllTexts(), StringComparison.Ordinal);
    }

    /// <summary>
    /// The read counterpart: a document filled from data, then walked and checked the way a consumer
    /// verifying its own output would.
    /// </summary>
    /// <remarks>
    /// This replaces a test that depended on a Word-produced file kept outside the repository, and was
    /// therefore permanently skipped. Reading foreign markup is covered instead by
    /// <see cref="TestKit.ForeignDocuments"/>, which builds the run splitting that made real files
    /// interesting in the first place without needing a binary fixture.
    /// </remarks>
    [Fact]
    public void AuthorFromData_ThenReadItBack_RoundTripsEveryValue()
    {
        (string Code, string Opened, string Closed)[] incidents =
        [
            ("INC-1", "2026-07-02", "2026-07-03"),
            ("INC-2", "2026-07-19", "2026-07-21"),
            ("INC-3", "2026-07-24", "2026-07-25"),
        ];

        using var stream = WriteAndValidate(body =>
        {
            body.AddHeading("Incidents", 1);

            var table = body.AddTable(
                incidents
                    .Select(incident => new[] { incident.Code, incident.Opened, incident.Closed })
                    .Prepend(["Code", "Opened", "Closed"]),
                new TableFormat { Borders = TableBorders.All });

            table.Rows[0].RepeatAsHeader();
        });

        using var document = OpenDocument(stream, isEditable: false);
        var read = document.GetBody().Tables[0];

        Assert.Equal(incidents.Length + 1, read.Rows.Count);
        Assert.True(read.Rows[0].IsHeader);

        for (var index = 0; index < incidents.Length; index++)
        {
            var row = read.Rows[index + 1];

            Assert.Equal(incidents[index].Code, row.Cells[0].GetAllTexts());
            Assert.Equal(incidents[index].Opened, row.Cells[1].GetAllTexts());
            Assert.Equal(incidents[index].Closed, row.Cells[2].GetAllTexts());
        }
    }
}
