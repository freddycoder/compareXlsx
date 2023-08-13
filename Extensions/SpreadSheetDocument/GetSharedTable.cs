using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace compareXlsx;

public static class GetSharedTableExtension {
    public static SharedStringTable? GetSharedTable(this SpreadsheetDocument doc) {
        var table = doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;

        return table;
    }
}