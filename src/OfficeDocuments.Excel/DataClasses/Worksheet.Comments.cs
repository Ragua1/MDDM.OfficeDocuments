namespace OfficeDocuments.Excel.DataClasses;

internal partial class Worksheet
{
    private CommentWriter? _commentWriter;

    private CommentWriter CommentWriter => _commentWriter ??= new CommentWriter(WorksheetPart, WorksheetElement);

    internal void SetCellComment(Cell cell, string text, string? author) => CommentWriter.Set(cell.CellReference, text, author);

    internal string? GetCellComment(string cellReference) => CommentWriter.Get(cellReference);
}
