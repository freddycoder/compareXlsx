using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public static class Extension {
    public static string GetCellValue(this Cell cell, SharedStringTable? table) {
        if (table != null && cell.DataType != null && cell.DataType == CellValues.SharedString) {
            return table.ElementAt(int.Parse(cell.InnerText)).InnerText;
        }

        return cell.InnerText;
    }
    public static string GetCellValue(this IEnumerator<Cell> cell, SharedStringTable table) {
        return cell.Current.GetCellValue(table);
    }
    public static SharedStringTable? GetSharedTable(this SpreadsheetDocument doc) {
        var table = doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;

        return table;
    }
}