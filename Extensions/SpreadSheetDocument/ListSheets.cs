using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace compareXlsx;

public static class ListSheetsExtension
{
    public static IEnumerable<(string, Worksheet)> ListSheets(this SpreadsheetDocument doc)
    {
        if (doc.WorkbookPart == null)
        {
            throw new ArgumentException("speadsheetDocument.Workbookpart must not be null");
        }

        // Get the workbook
        WorkbookPart wbPart = doc.WorkbookPart;

        // Get the sheets
        Sheets? sheets = (wbPart.Workbook?.Sheets) ?? throw new InvalidOperationException("sheets must be defined");

        // Loop through the sheets
        foreach (Sheet sheet in sheets.Cast<Sheet>())
        {
            if (sheet == null)
            {
                throw new InvalidOperationException("sheet is null");
            }

#nullable disable
            if (sheet.Id.HasValue == false)
            {
                throw new InvalidOperationException("sheet.Id has no value");
            }

            WorksheetPart wsPath =
                (WorksheetPart)wbPart.GetPartById(sheet.Id.Value);

            Worksheet ws = wsPath.Worksheet;

            //logger.Information("Returning sheet " + sheet.Name);

            yield return (sheet.Name, ws);
#nullable enable
        }
    }
}