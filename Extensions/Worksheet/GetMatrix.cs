using DocumentFormat.OpenXml.Spreadsheet;

namespace compareXlsx;

public static class GetMatrixExtension
{
    public static List<List<string>> GetMatrix(this Worksheet ws, SharedStringTable? table)
    {
        var rows = ws.Descendants<Row>().ToArray();

        //logger.Information("Sheet number of rows: " + rows.Length);

        var m = new List<List<string>>();

        foreach (Row row in rows)
        {
            m.Add(new List<string>());

            foreach (Cell cell in row.Descendants<Cell>())
            {
                m[^1].Add(cell.GetCellValue(table));
            }
        }

        return m;
    }
}