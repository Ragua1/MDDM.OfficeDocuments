using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.Extensions;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.DataClasses;

/// <summary>
/// Owns worksheet cell comments and the legacy VML drawing Excel uses to render them.
/// </summary>
internal sealed class CommentWriter(WorksheetPart worksheetPart, SpreadsheetLib.Worksheet worksheetElement)
{
    public void Set(string cellReference, string text, string? author)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            throw new ArgumentException("Comment text cannot be null or empty.", nameof(text));
        }

        XmlText.EnsureRepresentable(text, nameof(text), $"The comment on cell '{cellReference}'");
        if (author != null)
        {
            XmlText.EnsureRepresentable(author, nameof(author), $"The comment author for cell '{cellReference}'");
        }

        var commentsPart = worksheetPart.WorksheetCommentsPart ?? worksheetPart.AddNewPart<WorksheetCommentsPart>();
        var comments = commentsPart.Comments ??= new SpreadsheetLib.Comments(new SpreadsheetLib.Authors(), new SpreadsheetLib.CommentList());
        var authors = comments.Authors ?? comments.AppendChild(new SpreadsheetLib.Authors());
        var commentList = comments.CommentList ?? comments.AppendChild(new SpreadsheetLib.CommentList());

        author ??= "OfficeDocuments";
        var authorIndex = authors.Elements<SpreadsheetLib.Author>()
            .Select((item, index) => new { item, index })
            .FirstOrDefault(item => string.Equals(item.item.Text, author, StringComparison.Ordinal))?.index;

        if (authorIndex == null)
        {
            authors.Append(new SpreadsheetLib.Author(author));
            authorIndex = authors.Count() - 1;
        }

        var comment = commentList.Elements<SpreadsheetLib.Comment>()
            .FirstOrDefault(item => item.Reference?.Value == cellReference);

        if (comment == null)
        {
            comment = new SpreadsheetLib.Comment
            {
                Reference = cellReference,
                AuthorId = Convert.ToUInt32(authorIndex.Value)
            };
            commentList.Append(comment);
        }
        else
        {
            comment.AuthorId = Convert.ToUInt32(authorIndex.Value);
            comment.RemoveAllChildren<SpreadsheetLib.CommentText>();
        }

        comment.Append(
            new SpreadsheetLib.CommentText(
                new SpreadsheetLib.Run(
                    new SpreadsheetLib.RunProperties(),
                    new SpreadsheetLib.Text(text) { Space = SpaceProcessingModeValues.Preserve }
                )
            )
        );

        comments.Save();
        UpdateVml(commentList.Elements<SpreadsheetLib.Comment>().Select(item => item.Reference?.Value).OfType<string>());
    }

    public string? Get(string cellReference)
    {
        return worksheetPart.WorksheetCommentsPart?.Comments?
            .CommentList?
            .Elements<SpreadsheetLib.Comment>()
            .FirstOrDefault(comment => comment.Reference?.Value == cellReference)?
            .CommentText?
            .InnerText;
    }

    private void UpdateVml(IEnumerable<string> references)
    {
        var vmlPart = worksheetPart.VmlDrawingParts.FirstOrDefault() ?? worksheetPart.AddNewPart<VmlDrawingPart>();
        var vmlRelationshipId = worksheetPart.GetIdOfPart(vmlPart);

        var legacyDrawing = worksheetElement.GetFirstChild<SpreadsheetLib.LegacyDrawing>();
        if (legacyDrawing == null)
        {
            worksheetElement.Append(new SpreadsheetLib.LegacyDrawing { Id = vmlRelationshipId });
        }
        else
        {
            legacyDrawing.Id = vmlRelationshipId;
        }

        using var stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, Encoding.UTF8);
        writer.Write(BuildVml(references));
    }

    private static string BuildVml(IEnumerable<string> references)
    {
        var builder = new StringBuilder();
        builder.AppendLine("""<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">""");
        builder.AppendLine("""<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>""");
        builder.AppendLine("""<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>""");

        var shapeId = 1025;
        foreach (var reference in references)
        {
            var (rowIndex, columnIndex) = reference.GetExcelCellIndex();
            var zeroBasedRow = rowIndex - 1;
            var zeroBasedColumn = columnIndex - 1;
            var anchor = $"{zeroBasedColumn}, 15, {zeroBasedRow}, 2, {zeroBasedColumn + 3}, 15, {zeroBasedRow + 4}, 4";

            builder.AppendLine(
                $"""<v:shape id="_x0000_s{shapeId++}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{zeroBasedRow}</x:Row><x:Column>{zeroBasedColumn}</x:Column></x:ClientData></v:shape>"""
            );
        }

        builder.AppendLine("</xml>");
        return builder.ToString();
    }
}
