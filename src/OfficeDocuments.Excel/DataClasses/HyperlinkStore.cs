using DocumentFormat.OpenXml.Packaging;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.DataClasses;

/// <summary>
/// Owns worksheet cell hyperlinks and their part relationships. The caller is responsible for
/// any display text on the target cell; this store only manages the hyperlink XML.
/// </summary>
internal sealed class HyperlinkStore(WorksheetPart worksheetPart, SpreadsheetLib.Worksheet worksheetElement, WorksheetElementOrderer orderer)
{
    public void Set(string cellReference, string target)
    {
        var worksheetHyperlinks = worksheetElement.GetFirstChild<SpreadsheetLib.Hyperlinks>();
        if (worksheetHyperlinks == null)
        {
            worksheetHyperlinks = new SpreadsheetLib.Hyperlinks();
            orderer.InsertHyperlinks(worksheetHyperlinks);
        }

        var existingHyperlink = worksheetHyperlinks.Elements<SpreadsheetLib.Hyperlink>()
            .FirstOrDefault(hyperlink => hyperlink.Reference?.Value == cellReference);

        if (existingHyperlink?.Id?.Value is { Length: > 0 } existingRelationshipId)
        {
            worksheetPart.DeleteReferenceRelationship(existingRelationshipId);
        }

        existingHyperlink?.Remove();

        SpreadsheetLib.Hyperlink hyperlink;
        if (Uri.TryCreate(target, UriKind.Absolute, out var absoluteUri))
        {
            var relationship = worksheetPart.AddHyperlinkRelationship(absoluteUri, true);
            hyperlink = new SpreadsheetLib.Hyperlink
            {
                Reference = cellReference,
                Id = relationship.Id
            };
        }
        else
        {
            hyperlink = new SpreadsheetLib.Hyperlink
            {
                Reference = cellReference,
                Location = target.TrimStart('#')
            };
        }

        worksheetHyperlinks.Append(hyperlink);
    }

    public string? Get(string cellReference)
    {
        var hyperlink = worksheetElement.GetFirstChild<SpreadsheetLib.Hyperlinks>()?
            .Elements<SpreadsheetLib.Hyperlink>()
            .FirstOrDefault(current => current.Reference?.Value == cellReference);

        if (hyperlink == null)
        {
            return null;
        }

        if (hyperlink.Id?.Value is { Length: > 0 } relationshipId)
        {
            return worksheetPart.HyperlinkRelationships.FirstOrDefault(relationship => relationship.Id == relationshipId)?.Uri.ToString();
        }

        return hyperlink.Location?.Value;
    }
}
