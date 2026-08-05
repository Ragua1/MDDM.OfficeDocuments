using OfficeDocuments.Excel.Enums;

namespace OfficeDocuments.Excel.DataClasses;

internal partial class Worksheet
{
    private WorksheetImageWriter? _imageWriter;

    private WorksheetImageWriter ImageWriter => _imageWriter ??= new WorksheetImageWriter(WorksheetPart, WorksheetElement);

    public void AddImage(string filePath, uint fromColumn, uint fromRow, uint toColumn, uint toRow) =>
        ImageWriter.AddImage(filePath, fromColumn, fromRow, toColumn, toRow);

    public void AddImage(Stream imageStream, ImageType imageType, uint fromColumn, uint fromRow, uint toColumn, uint toRow) =>
        ImageWriter.AddImage(imageStream, imageType, fromColumn, fromRow, toColumn, toRow);
}
