using OfficeDocuments.Excel.DataClasses;

namespace OfficeDocuments.Excel;

public partial class Spreadsheet
{
    private WorkbookProtector? _workbookProtector;

    private WorkbookProtector WorkbookProtector => _workbookProtector ??= new WorkbookProtector(WorkbookPartInternal);

    public void ProtectWorkbook(string? password = null) => WorkbookProtector.Protect(password);
}
