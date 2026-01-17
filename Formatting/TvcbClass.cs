using System.Diagnostics;
using ExcelFormatterConsole.Utility;
using OfficeOpenXml;

namespace ExcelFormatterConsole.Formatting;

public class TvcbClass
// Total Volume Class Breakdown
{
    private static readonly Stopwatch Stopwatch = Stopwatch.StartNew();

    private static readonly Dictionary<string, string> DirectionTranslations = new (StringComparer.OrdinalIgnoreCase)
    {
        { "right", "doprava" },
        { "left", "dolava" },
        { "thru", "priamo" },
        { "u-turn", "otočenie" },

        { "hard right", "prudko doprava" },
        { "hard left", "prudko doľava" },
        { "slight right", "mierne doprava" },
        { "slight left", "mierne doľava" },
        { "bear right", "mierne doprava" },
        { "bear left", "mierne doľava" },

        { "app total", "spolu"},
        {"int total", "celkom"}
    };

    private static void TimedLog(string logMessage)
    {
        Stopwatch.Stop();
        Console.WriteLine($"[{Stopwatch.Elapsed.TotalMilliseconds} ms] ----- {logMessage} -----");
        Stopwatch.Restart();
    }

    public static ExcelWorksheet FindCorrectWorksheet(ExcelPackage genPackage)
    {
        var genWs = genPackage.Workbook.Worksheets.FirstOrDefault(ws => ws.Index > 4 && ws.Name.ToLower() == "total volume class breakdown");
        return genWs ?? throw new MissingWorksheetException($"TvcbClass.FindCorrectWorksheet | Failed to find correct worksheet. | Checked file name: '{genPackage.File.Name}'.");
    }

    public static ExcelWorksheet Prepare(ExcelPackage toFormatPackage)
    {
        return toFormatPackage.Workbook.Worksheets.Add("Celkové údaje 12hod");
    }

    public static void FormatMeasuredTime(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {
        var row = 4;
        while (row < 1000)
        {
            var cellValue = genWs.Cells["A" + row].Value ?? "";
            if (string.IsNullOrWhiteSpace(cellValue.ToString()))
            {
                row++;
                continue;
            }

            if (!DateTime.TryParse(cellValue.ToString(), out DateTime dtCellValue))
            {
                break;
            }

            row++;
            var nextCellValue = genWs.Cells["A" + row].Value ?? "";

            if (string.IsNullOrWhiteSpace(nextCellValue.ToString()) || !DateTime.TryParse(nextCellValue.ToString(), out DateTime dtNextCellValue))
            {
                row--;
                var beforeLastRow = row - 1;

                cellValue = genWs.Cells["A" + beforeLastRow].Value;
                nextCellValue = genWs.Cells["A" + row].Value;

                var difference = DateTime.Parse(nextCellValue.ToString() ?? "00:00") - DateTime.Parse(cellValue.ToString() ?? "");

                nextCellValue = DateTime.Parse(nextCellValue.ToString() ?? "00:00").AddMinutes(difference.TotalMinutes);
                cellValue = DateTime.Parse(nextCellValue.ToString() ?? "00:00").AddMinutes(-difference.TotalMinutes);

                toFormatWs.Cells["A" + row].Value = $"{cellValue:HH:mm} - {nextCellValue:HH:mm}";

                break;
            }

            var correctCellRow = row - 1;
            toFormatWs.Cells["A" + correctCellRow].Value = $"{dtCellValue:HH:mm} - {dtNextCellValue:HH:mm}";
        }

        TimedLog($"{toFormatWs.Name} | Applied date formatting.");
    }

    public static void Navigation(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {
        toFormatWs.Cells["A1"].Value = "Smer od";
        toFormatWs.Cells["A2"].Value = "Orientácia";
        toFormatWs.Cells["A3"].Value = "Čas";

        var lastColumn = genWs.Dimension.End.Column;
        var lastRow = genWs.Dimension.End.Row;
        var targetValue = 1;
        var column = 2;

        List<List<string>> listOfAllData = [];
        while (column <= 1000)
        {
            var cellValue = genWs.Cells[1, column].Value?.ToString() ?? "";
            if (!int.TryParse(cellValue.Split("-")[0].Trim(), out var intValue) || intValue != targetValue)
            {
                column++;
                continue;
            }

            targetValue++;

            List<string> setColumnData = [];
            for (var row = 1; row <= lastRow; row++)
            {
                var detectedKeywordAmount = 0;
                var detectedCheckWordAmount = 0;
                for (var innerColumn = column; innerColumn <= lastColumn; innerColumn++)
                {
                    switch (row)
                    {
                        case 1 or 2:
                            var checkWord = genWs.Cells[3, column].Value?.ToString() ?? "";
                            if (string.IsNullOrWhiteSpace(checkWord))
                            {
                                HelperFunctions.ErrorLog("checkWord is null or empty");
                                continue;
                            }

                            if (checkWord.ToLower() != "right" || detectedCheckWordAmount >= 1)
                            {
                                continue;
                            }

                            detectedCheckWordAmount++;

                            var directionName = genWs.Cells[row, column].Value?.ToString() ?? "";
                            if (string.IsNullOrWhiteSpace(directionName))
                            {
                                HelperFunctions.ErrorLog("direction name is null or empty");
                                continue;
                            }

                            var fullMergedRange = genWs.MergedCells[1, column];
                            toFormatWs.Cells[fullMergedRange].Merge = true;

                            setColumnData.Add(directionName);
                            break;
                        case 3:
                            break;
                        default:
                            var keyword = genWs.Cells[3, innerColumn].Value?.ToString() ?? "";
                            if (string.IsNullOrWhiteSpace(keyword))
                            {
                                HelperFunctions.ErrorLog("keyword value is null or empty");
                                continue;
                            }

                            if (keyword.ToLower() == "right")
                            {
                                detectedKeywordAmount++;
                            }
                            // not working correctly
                            else if (detectedKeywordAmount > 1)
                            {
                                detectedKeywordAmount = 0;
                                setColumnData.Add("columnEnd");
                                goto endOfRow;
                            }

                            cellValue = genWs.Cells[row, innerColumn].Value?.ToString() ?? "";
                            if (string.IsNullOrWhiteSpace(cellValue) || !double.TryParse(cellValue, out _))
                            {
                                HelperFunctions.ErrorLog($"cell value is null, empty or a non-numeric value | cellValue='{cellValue}'");
                                continue;
                            }

                            setColumnData.Add(cellValue);
                            break;
                    }
                }

                endOfRow: ;
            }

            listOfAllData.Add(setColumnData);
            column = 2;
        }
    }

    public static void SecondaryNavigation()
    {

    }

    public static void PrimaryDataReading()
    {

    }

    public static void PrimaryDataWriting()
    {

    }

    public static void Style(ExcelWorksheet toFormatWs)
    {
        toFormatWs.Cells.AutoFitColumns();
    }
}
