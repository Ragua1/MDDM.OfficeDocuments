namespace OfficeDocuments.Excel.DataClasses;

internal partial class Worksheet
{
    private WorksheetElementOrderer? _elementOrderer;

    private WorksheetElementOrderer ElementOrderer => _elementOrderer ??= new WorksheetElementOrderer(WorksheetElement, Element);
}
