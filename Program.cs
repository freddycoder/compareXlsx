using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

if (args.Length == 0 || args[0].ToLower() == "--help") {
    Console.WriteLine("compareXlsx.exe");
    Console.WriteLine("Author: Frédéric Jacques");
    Console.WriteLine();
    Console.WriteLine("definition:");
    Console.WriteLine(".\\compareXlsx.exe <file1> <file2> <throwException:bool> <ignoreRowNumber:int> <ignoreDateDiff:bool>");
    Console.WriteLine("");
}

var file1 = args[0];
var file2 = args[1];
var throwException = false;
int? ignoreRowOne = null;
bool ignoreDateDiff = false;

if (args.Length > 2) {
    throwException = args[2].ToLower() == "true";
}

if (args.Length > 3) {
    ignoreRowOne = int.Parse(args[3]);
}

if (args.Length > 4) {
    ignoreDateDiff = args[4].ToLower() == "true";
}

Console.WriteLine($"Compare {file1} to {file2}");

using SpreadsheetDocument sd1 = SpreadsheetDocument.Open(file1, false);
using SpreadsheetDocument sd2 = SpreadsheetDocument.Open(file2, false);

var feuilles1 = SeletionneFeuilles(sd1).ToArray();
var feuilles2 = SeletionneFeuilles(sd2).ToArray();

var table1 = sd1.GetSharedTable();
var table2 = sd2.GetSharedTable();

if (feuilles1.Length != feuilles2.Length) {
    Console.WriteLine($"Nombre de feuilles différentes. {feuilles1.Length} et {feuilles2.Length}");
    if (throwException) {
        throw new Exception($"Nombre de feuilles différentes. {feuilles1.Length} et {feuilles2.Length}");
    }
}

long celluleComparer = 0;
long differenceTotal = 0;

for (int i = 0; i < feuilles1.Length; i++) {
    var ws1 = feuilles1[i].Item2;
    var ws2 = feuilles2[i].Item2;

    var m1 = GetMatrix(table1, ws1);
    var m2 = GetMatrix(table2, ws2);

    if (m1.Count != m2.Count) {
        Console.WriteLine($"Matrices have different number of row for sheet {feuilles1[i].Item1} and {feuilles2[i].Item1}. ({m1.Count} and {m2.Count})");
        if (throwException) {
            throw new Exception($"Matrices have different number of row for sheet {feuilles1[i].Item1} and {feuilles2[i].Item1}. ({m1.Count} and {m2.Count})");
        }
    }

    for (int r = 0; r < m1.Count; r++) {
        if (r == ignoreRowOne) {
            continue;
        }

        var cells1 = m1[r];
        var cells2 = m2.Count > r ? m2[r] : null;

        if (cells1.Count != cells2?.Count) {
            Console.WriteLine($"Row {i} of sheets {feuilles1[i].Item1} and {feuilles2[i].Item1} and differents number of cells. ({cells1.Count} and {cells2?.Count})");
            if (throwException) {
                throw new Exception($"Row {i} of sheets {feuilles1[i].Item1} and {feuilles2[i].Item1} and differents number of cells. ({cells1.Count} and {cells2?.Count})");
            }
            differenceTotal++;
        }

        if (cells2 != null) {
            for (int c = 0; c < m1[r].Count; c++) {
                //Console.WriteLine($"Comparing {m1[r][c]} to {m2[r][c]}");
                celluleComparer++;

                if (m1[r][c] != m2[r][c]) {
                    if (ignoreDateDiff && 
                        DateTimeOffset.TryParse(m1[r][c], out _) && 
                        DateTimeOffset.TryParse(m2[r][c], out _)) {
                        continue;
                    }

                    differenceTotal++;
                    Console.WriteLine($"Differences in sheet {feuilles1[i].Item1} and {feuilles2[i].Item1}. Row {r} Column {c}. Values are {m1[r][c]} and {m2[r][c]}");
                    if (throwException) {
                        throw new Exception($"Differences in sheet {feuilles1[i].Item1} and {feuilles2[i].Item1}. Row {r} Column {c}. Values are {m1[r][c]} and {m2[r][c]}");
                    }
                }
            }
        }
    } 
}

Console.WriteLine($"{celluleComparer} cellules comparées");
Console.WriteLine($"{differenceTotal} différences trouvées");
Console.WriteLine("Traitement terminé.");

static IEnumerable<(string, Worksheet)> SeletionneFeuilles(SpreadsheetDocument spreadsheetDocument)
{
    if (spreadsheetDocument.WorkbookPart == null)
    {
        throw new ArgumentException("speadsheetDocument.Workbookpart must not be null");
    }

    // Get the workbook
    WorkbookPart wbPart = spreadsheetDocument.WorkbookPart;

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

        //Console.WriteLine("Returning sheet " + sheet.Name);

        yield return (sheet.Name, ws);
#nullable enable
    }
}

List<List<string>> GetMatrix(SharedStringTable? table, Worksheet ws)
{
    var rows = ws.Descendants<Row>().ToArray();

    //Console.WriteLine("Sheet number of rows: " + rows.Length);

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