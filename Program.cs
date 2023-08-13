using DocumentFormat.OpenXml.Packaging;

namespace compareXlsx;

class Program
{
    public static void Main(string[] args)
    {
        var (provider, logger, option) = Setup.GetAppServices(args);

        if (args.Length == 0 || args[0].ToLower() == "--help")
        {
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

        if (args.Length > 2)
        {
            throwException = args[2].ToLower() == "true";
        }

        if (args.Length > 3)
        {
            ignoreRowOne = int.Parse(args[3]);
        }

        if (args.Length > 4)
        {
            ignoreDateDiff = args[4].ToLower() == "true";
        }

        logger.Information($"Compare {file1} to {file2}");

        using SpreadsheetDocument sd1 = SpreadsheetDocument.Open(file1, false);
        using SpreadsheetDocument sd2 = SpreadsheetDocument.Open(file2, false);

        var feuilles1 = sd1.ListSheets().ToArray();
        var feuilles2 = sd2.ListSheets().ToArray();

        var table1 = sd1.GetSharedTable();
        var table2 = sd2.GetSharedTable();

        if (feuilles1.Length != feuilles2.Length)
        {
            logger.Warning($"Nombre de feuilles différentes. {feuilles1.Length} et {feuilles2.Length}");
            if (throwException)
            {
                throw new Exception($"Nombre de feuilles différentes. {feuilles1.Length} et {feuilles2.Length}");
            }
        }

        long celluleComparer = 0;
        long differenceTotal = 0;

        for (int i = 0; i < feuilles1.Length; i++)
        {
            var ws1 = feuilles1[i].Item2;
            var ws2 = feuilles2[i].Item2;

            var m1 = ws1.GetMatrix(table1);
            var m2 = ws2.GetMatrix(table2);

            if (m1.Count != m2.Count)
            {
                logger.Warning($"Matrices have different number of row for sheet {feuilles1[i].Item1} and {feuilles2[i].Item1}. ({m1.Count} and {m2.Count})");
                if (throwException)
                {
                    throw new Exception($"Matrices have different number of row for sheet {feuilles1[i].Item1} and {feuilles2[i].Item1}. ({m1.Count} and {m2.Count})");
                }
            }

            for (int r = 0; r < m1.Count; r++)
            {
                if (r == ignoreRowOne)
                {
                    continue;
                }

                var cells1 = m1[r];
                var cells2 = m2.Count > r ? m2[r] : null;

                if (cells1.Count != cells2?.Count)
                {
                    logger.Warning($"Row {i} of sheets {feuilles1[i].Item1} and {feuilles2[i].Item1} and differents number of cells. ({cells1.Count} and {cells2?.Count})");
                    if (throwException)
                    {
                        throw new Exception($"Row {i} of sheets {feuilles1[i].Item1} and {feuilles2[i].Item1} and differents number of cells. ({cells1.Count} and {cells2?.Count})");
                    }
                    differenceTotal++;
                }

                if (cells2 != null)
                {
                    for (int c = 0; c < m1[r].Count; c++)
                    {
                        //logger.Information($"Comparing {m1[r][c]} to {m2[r][c]}");
                        celluleComparer++;

                        if (m1[r][c] != m2[r][c])
                        {
                            if (ignoreDateDiff &&
                                DateTimeOffset.TryParse(m1[r][c], out _) &&
                                DateTimeOffset.TryParse(m2[r][c], out _))
                            {
                                continue;
                            }

                            differenceTotal++;
                            logger.Warning($"Differences in sheet {feuilles1[i].Item1} and {feuilles2[i].Item1}. Row {r} Column {c}. Values are {m1[r][c]} and {m2[r][c]}");
                            if (throwException)
                            {
                                throw new Exception($"Differences in sheet {feuilles1[i].Item1} and {feuilles2[i].Item1}. Row {r} Column {c}. Values are {m1[r][c]} and {m2[r][c]}");
                            }
                        }
                    }
                }
            }
        }

        logger.Information($"{celluleComparer} cellules comparées");
        logger.Information($"{differenceTotal} différences trouvées");
        logger.Information("Traitement terminé.");
    }
}