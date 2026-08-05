using OfficeDocuments.Excel.DataClasses;
using OfficeDocuments.Excel.Interfaces;

namespace OfficeDocuments.Excel;

public partial class Spreadsheet
{
    private NamedRangeManager? _namedRangeManager;

    private NamedRangeManager NamedRangeManager =>
        _namedRangeManager ??= new NamedRangeManager(WorkbookPartInternal, name => GetSheetIndex(GetSheet(GetWorksheetOrThrow(name))));

    public void AddNamedRange(string name, IRange range, bool worksheetScoped = false) =>
        NamedRangeManager.Add(name, range, worksheetScoped);
}
