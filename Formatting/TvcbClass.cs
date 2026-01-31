using System.Diagnostics;
using ExcelFormatterConsole.Utility;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelFormatterConsole.Formatting;

public class TvcbClass
// Total Volume Class Breakdown
{
    private static readonly Stopwatch Stopwatch = Stopwatch.StartNew();

    private static void TimedLog(string logMessage)
    {
        Stopwatch.Stop();
        Console.WriteLine($"[{Stopwatch.Elapsed.TotalMilliseconds} ms] ----- {logMessage} -----");
        Stopwatch.Restart();
    }

    private static readonly Dictionary<string, string> VehicleCategoryTranslations = new(StringComparer.OrdinalIgnoreCase)
    {
        {"motorcycles", "M"},
        {"lights", "LV"},
        {"single-unit trucks", "NV"},
        {"articulated trucks", "TNV"},
        {"buses", "A"},
        {"bicycles on road", "B"},
        {"articulated buses", "AK"},
        {"pedestrians", "CH"},
    };

    public static ExcelWorksheet FindCorrectWorksheet(ExcelPackage genPackage)
    {
        var genWs = genPackage.Workbook.Worksheets.FirstOrDefault(ws => ws.Index >= 4 && ws.Name.ToLower() == "total volume class breakdown");
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

    public static List<List<string>> PrimaryDataReading(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
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
            List<string> lastColumnData = [];

            var dataColumnRange = 0;
            for (var row = 1; row <= lastRow; row++)
            {
                var detectedCheckWordAmount = 0;
                for (var innerColumn = column; innerColumn <= lastColumn; innerColumn++)
                {
                    switch (row)
                    {
                        case 1:
                            if (innerColumn != column) continue;

                            var directionName = genWs.Cells[row, column].Value?.ToString() ?? "";
                            if (string.IsNullOrWhiteSpace(directionName))
                            {
                                HelperFunctions.WarningLog("direction name is null or empty");
                                continue;
                            }

                            var fullMergedRange = genWs.MergedCells[1, column];
                            toFormatWs.Cells[fullMergedRange].Merge = true;
                            fullMergedRange = genWs.MergedCells[2, column];
                            toFormatWs.Cells[fullMergedRange].Merge = true;

                            dataColumnRange = new ExcelAddress(fullMergedRange).Columns;
                            setColumnData.Add(directionName);
                            break;
                        case 3 or 2:
                            continue;
                        default:
                            var keyword = genWs.Cells[3, innerColumn].Value?.ToString() ?? "";
                            if (string.IsNullOrWhiteSpace(keyword))
                            {
                                HelperFunctions.WarningLog("keyword value is null or empty");
                                continue;
                            }

                            if (innerColumn >= column + dataColumnRange)
                            {
                                setColumnData.Add("columnEnd");
                                goto endOfRow;
                            }

                            cellValue = genWs.Cells[row, innerColumn].Value?.ToString() ?? "";
                            if (string.IsNullOrWhiteSpace(cellValue) || !double.TryParse(cellValue, out _))
                            {
                                HelperFunctions.ErrorLog($"cell value is null, empty or a non-numeric value | cellValue='{cellValue}'");
                            }

                            if (keyword.ToLower() == "int total")
                            {
                                setColumnData.Add("columnEnd");
                                lastColumnData.Add(cellValue);
                                goto endOfRow;
                            }

                            try
                            {
                                Console.WriteLine($"cellValue={cellValue} row={row} column={innerColumn} direction={setColumnData[0]}");
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e);
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

        return listOfAllData;
    }

    public static void PrimaryDataWriting(ExcelWorksheet genWs, ExcelWorksheet toFormatWs, List<List<string>> listsOfAllData)
    {
        var lastRow = genWs.Dimension.End.Row;
        var lastColumn = genWs.Dimension.End.Column;

        var spaceBetweenDataColumn = 2;
        var leftNumber = 1;

        foreach (var dataList in listsOfAllData)
        {
            var listIndex = 0;
            for (var row = 1; row <= lastRow; row++)
            {
                switch (row)
                {
                    case 2:
                        continue;
                    case 3:
                        var mergedRange = toFormatWs.MergedCells[1, 2];
                        var address = new ExcelAddress(mergedRange);


                        var rightNumber = leftNumber + 1;
                        var totalDirections = address.Columns;

                        var dataColumnRange = totalDirections;
                        for (var column = 2; column <= dataColumnRange + 1; column++)
                        {
                            var genData = genWs.Cells[row, column].Value?.ToString() ?? "";
                            if (genData.Contains("peds", StringComparison.InvariantCultureIgnoreCase)) totalDirections--;
                        }

                        for (var column = spaceBetweenDataColumn; column <= lastColumn; column++)
                        {
                            if (rightNumber >= totalDirections)
                            {
                                rightNumber = 1;
                            }

                            if (leftNumber == rightNumber)
                            {
                                toFormatWs.Cells[row, column].Value = $"{leftNumber} - {rightNumber}";
                                toFormatWs.Cells[row, column + 1].Value = "Spolu";

                                var cellValue = toFormatWs.Cells[row + 1, column + 2].Value?.ToString() ?? "";
                                var cellValueAhead = toFormatWs.Cells[row + 1, column + 3].Value?.ToString() ?? "";
                                if (double.TryParse(cellValue, out _) || double.TryParse(cellValueAhead, out _)) break;

                                toFormatWs.Cells[row, column + 2].Value = "Peds CW";
                                toFormatWs.Cells[row, column + 3].Value = "Peds CCW";

                                break;
                            }

                            toFormatWs.Cells[row, column].Value = $"{leftNumber} - {rightNumber}";

                            rightNumber++;
                        }

                        leftNumber++;
                        break;
                    default:
                        for (var column = spaceBetweenDataColumn; column <= lastColumn; column++, listIndex++)
                        {
                            var data = dataList[listIndex];
                            if (data == "columnEnd")
                            {
                                goto endOfRow;
                            }

                            if (string.IsNullOrWhiteSpace(data))
                            {
                                data = string.Empty;
                            }

                            if (!decimal.TryParse(data ,out var decimalData) && row > 3)
                            {
                                HelperFunctions.WarningLog($"Couldn't parse into a decimal value | row='{row}' column='{column}'");
                            }

                            toFormatWs.Cells[row, column].Value = decimalData == 0 ? data : decimalData;
                            if (row is 1)
                            {
                                goto endOfRow;
                            }

                            if (row != lastRow) continue;
                            spaceBetweenDataColumn++;
                        }
                        break;
                }

                endOfRow:
                if (row != 3) listIndex++;
            }
        }
    }

    public static void SecondaryNavigation(ExcelWorksheet toFormatWs, ExcelWorksheet genWs)
    {
        Dictionary<string, string> secondaryNavigationCategories = new(StringComparer.OrdinalIgnoreCase)
        {
            { "grand total", "spolu" },
            { "% approach", "% pomer na smer" },
            { "% total", "% celkový pomer" },
        };

        var lastRow = genWs.Dimension.End.Row;
        for (var row = 1; row <= lastRow; row++)
        {
            var cellValue = genWs.Cells[row, 1].Value?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(cellValue))
                continue;

            if (secondaryNavigationCategories.TryGetValue(cellValue, out var secondaryTranslation))
            {
                toFormatWs.Cells[row, 1].Value = secondaryTranslation;
            }

            if (!VehicleCategoryTranslations.TryGetValue(cellValue, out var translation)) continue;
            toFormatWs.Cells[row, 1].Value = translation;
            toFormatWs.Cells[row + 1, 1].Value = $"% {translation}";
        }
    }


    public static void Style(ExcelWorksheet toFormatWs)
    {
        toFormatWs.Cells["A:A"].AutoFitColumns();
        // toFormatWs.Cells["B:XX"].Style.Numberformat.Format = "0.#%"; <-- also needs work
        // toFormatWs.Cells["B:XX"].AutoFitColumns(); <-- needs some work
        toFormatWs.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        toFormatWs.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

    }
}
