using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drw = DocumentFormat.OpenXml.Drawing;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;
using XdrSpr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeDocuments.Excel.Advanced;

/// <summary>
/// Owns worksheet image embedding: the drawings part, the two-cell anchor, and the
/// worksheet-level drawing reference.
/// </summary>
internal sealed class WorksheetImageWriter(WorksheetPart worksheetPart, SpreadsheetLib.Worksheet worksheetElement)
{
    public void AddImage(string filePath, uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("Image file not found.", filePath);
        }

        var imageType = DetectImageType(filePath);
        using var stream = File.OpenRead(filePath);
        AddImage(stream, imageType, fromColumn, fromRow, toColumn, toRow);
    }

    public void AddImage(Stream imageStream, ImageType imageType, uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        ArgumentNullException.ThrowIfNull(imageStream);
        if (fromColumn < 1)
        {
            throw new ArgumentException("fromColumn must be at least 1.", nameof(fromColumn));
        }

        if (fromRow < 1)
        {
            throw new ArgumentException("fromRow must be at least 1.", nameof(fromRow));
        }

        if (toColumn < fromColumn)
        {
            throw new ArgumentException("toColumn must be greater than or equal to fromColumn.", nameof(toColumn));
        }

        if (toRow < fromRow)
        {
            throw new ArgumentException("toRow must be greater than or equal to fromRow.", nameof(toRow));
        }

        var drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
        var imagePart = drawingsPart.AddImagePart(ToImagePartType(imageType));
        imagePart.FeedData(imageStream);
        var imageRelId = drawingsPart.GetIdOfPart(imagePart);

        drawingsPart.WorksheetDrawing ??= new XdrSpr.WorksheetDrawing();
        var existingCount = drawingsPart.WorksheetDrawing.Elements<XdrSpr.TwoCellAnchor>().Count()
            + drawingsPart.WorksheetDrawing.Elements<XdrSpr.OneCellAnchor>().Count();
        var pictureId = (uint)(existingCount + 1);

        drawingsPart.WorksheetDrawing.Append(BuildTwoCellAnchor(imageRelId, pictureId, fromColumn, fromRow, toColumn, toRow));
        EnsureDrawingElement(worksheetPart.GetIdOfPart(drawingsPart));
    }

    private static XdrSpr.TwoCellAnchor BuildTwoCellAnchor(
        string imageRelId, uint pictureId, uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        var anchor = new XdrSpr.TwoCellAnchor();
        anchor.Append(new XdrSpr.FromMarker(
            new XdrSpr.ColumnId((fromColumn - 1).ToString()),
            new XdrSpr.ColumnOffset("0"),
            new XdrSpr.RowId((fromRow - 1).ToString()),
            new XdrSpr.RowOffset("0")
        ));
        anchor.Append(new XdrSpr.ToMarker(
            new XdrSpr.ColumnId(toColumn.ToString()),
            new XdrSpr.ColumnOffset("0"),
            new XdrSpr.RowId(toRow.ToString()),
            new XdrSpr.RowOffset("0")
        ));
        var picture = new XdrSpr.Picture();
        picture.Append(new XdrSpr.NonVisualPictureProperties(
            new XdrSpr.NonVisualDrawingProperties { Id = pictureId, Name = $"Image{pictureId}" },
            new XdrSpr.NonVisualPictureDrawingProperties()
        ));
        picture.Append(new XdrSpr.BlipFill(
            new Drw.Blip { Embed = imageRelId },
            new Drw.Stretch(new Drw.FillRectangle())
        ));
        picture.Append(new XdrSpr.ShapeProperties(
            new Drw.Transform2D(
                new Drw.Offset { X = 0, Y = 0 },
                new Drw.Extents { Cx = 0, Cy = 0 }
            ),
            new Drw.PresetGeometry(new Drw.AdjustValueList()) { Preset = Drw.ShapeTypeValues.Rectangle }
        ));
        anchor.Append(picture);
        anchor.Append(new XdrSpr.ClientData());
        return anchor;
    }

    private void EnsureDrawingElement(string drawingRelId)
    {
        if (worksheetElement.GetFirstChild<SpreadsheetLib.Drawing>() != null)
        {
            return;
        }

        var drawing = new SpreadsheetLib.Drawing { Id = drawingRelId };
        var legacyDrawing = worksheetElement.GetFirstChild<SpreadsheetLib.LegacyDrawing>();
        if (legacyDrawing != null)
        {
            worksheetElement.InsertBefore(drawing, legacyDrawing);
            return;
        }

        var tableParts = worksheetElement.GetFirstChild<SpreadsheetLib.TableParts>();
        if (tableParts != null)
        {
            worksheetElement.InsertBefore(drawing, tableParts);
            return;
        }

        worksheetElement.AppendChild(drawing);
    }

    private static PartTypeInfo ToImagePartType(ImageType imageType) => imageType switch
    {
        ImageType.Png => ImagePartType.Png,
        ImageType.Jpeg => ImagePartType.Jpeg,
        ImageType.Gif => ImagePartType.Gif,
        ImageType.Bmp => ImagePartType.Bmp,
        ImageType.Tiff => ImagePartType.Tiff,
        _ => throw new ArgumentOutOfRangeException(nameof(imageType))
    };

    private static ImageType DetectImageType(string filePath)
    {
        return Path.GetExtension(filePath).ToLowerInvariant() switch
        {
            ".png" => ImageType.Png,
            ".jpg" or ".jpeg" => ImageType.Jpeg,
            ".gif" => ImageType.Gif,
            ".bmp" => ImageType.Bmp,
            ".tiff" or ".tif" => ImageType.Tiff,
            var ext => throw new ArgumentException($"Unsupported image format: '{ext}'.", nameof(filePath))
        };
    }
}
