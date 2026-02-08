using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using DocumentFormat.OpenXml.Presentation;
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
        { "app total", "spolu" }
    };

    private static readonly Dictionary<string, string> WorldDirectionTranslations = new(StringComparer.OrdinalIgnoreCase)
    {
        ["north"] = "sever",
        ["north-northeast"] = "sever-severovýchod",
        ["northeast"] = "severovýchod",
        ["east-northeast"] = "východ-severovýchod",
        ["east"] = "východ",
        ["east-southeast"] = "východ-juhovýchod",
        ["southeast"] = "juhovýchod",
        ["south-southeast"] = "juh-juhovýchod",
        ["south"] = "juh",
        ["south-southwest"] = "juh-juhozápad",
        ["southwest"] = "juhozápad",
        ["west-southwest"] = "západ-juhozápad",
        ["west"] = "západ",
        ["west-northwest"] = "západ-severozápad",
        ["northwest"] = "severozápad",
        ["north-northwest"] = "sever-severozápad"
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

    public static (List<List<string>>, List<int>, List<string>) PrimaryDataReading(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {
        toFormatWs.Cells["A1"].Value = "Smer od";
        toFormatWs.Cells["A2"].Value = "Orientácia";
        toFormatWs.Cells["A3"].Value = "Čas";

        var lastColumn = genWs.Dimension.End.Column;
        var lastRow = genWs.Dimension.End.Row;
        var targetValue = 1;
        var column = 2;

        List<List<string>> listOfAllData = [];
        List<int> directionColumns = [];
        List<string> lastColumnData = [];
        List<string> allDirections = [];

        while (column <= 1000)
        {
            var cellValue = genWs.Cells[1, column].Value?.ToString() ?? "";
            if (!int.TryParse(cellValue.Split("-")[0].Trim(), out var intValue) || intValue != targetValue)
            {
                column++;
                continue;
            }

            targetValue++;
            directionColumns.Add(column);

            List<string> setColumnData = [];
            var dataColumnRange = 0;

            for (var row = 1; row <= lastRow; row++)
            {
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

                            dataColumnRange = new ExcelAddress(genWs.MergedCells[1, column]).Columns;
                            setColumnData.Add(directionName);
                            break;
                        case 3:
                            if (innerColumn < column + dataColumnRange)
                            {
                                cellValue = genWs.Cells[row, innerColumn].Value?.ToString() ?? "";
                                if (string.IsNullOrWhiteSpace(cellValue))
                                {
                                    HelperFunctions.WarningLog("Failed to get direction");
                                    cellValue = "unknown";
                                }

                                if (cellValue.Contains("int total")) continue;

                                allDirections.Add(cellValue);
                            }

                            break;
                        default:
                            if (row == 2)
                            {
                                cellValue = genWs.Cells[row, column].Value?.ToString() ?? "";

                                if (setColumnData.Contains(cellValue))
                                {
                                    continue;
                                }
                            }
                            else
                            {
                                cellValue = genWs.Cells[row, innerColumn].Value?.ToString() ?? "";
                            }

                            if (string.IsNullOrWhiteSpace(cellValue) || !double.TryParse(cellValue, out _))
                            {
                                HelperFunctions.WarningLog($"cell value is null, empty or a non-numeric value | cellValue='{cellValue}'");
                            }

                            if (innerColumn >= column + dataColumnRange)
                            {
                                setColumnData.Add("columnEnd");
                                goto endOfRow;
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

        for (var row = 4; row <= lastRow; row++)
        {
            var cellValue = genWs.Cells[row, lastColumn].Value?.ToString() ?? "";
            lastColumnData.Add(cellValue);
        }

        listOfAllData.Add(lastColumnData);
        return (listOfAllData, directionColumns, allDirections);
    }

    public static void PrimaryDataWriting(ExcelWorksheet genWs, ExcelWorksheet toFormatWs, List<List<string>> listsOfAllData, List<int> directionColumns, List<string> allDirections)
    {
        var lastRow = genWs.Dimension.End.Row;
        var lastColumn = genWs.Dimension.End.Column;

        var spaceBetweenDataColumn = 2;

        toFormatWs.Cells[1, lastColumn].Value = "Celkom";
        toFormatWs.Cells[1, lastColumn, 3, lastColumn].Merge = true;

        var directionColumnListIndex = 0;
        foreach (var dataList in listsOfAllData)
        {
            var listIndex = 0;
            if (dataList.Equals(listsOfAllData.Last()))
            {
                for (var row = 4; row <= lastRow; row++)
                {
                    toFormatWs.Cells[row, lastColumn].Value = dataList[listIndex++];
                }
                break;
            }

            var totalDirections = 0;
            for (var row = 1; row <= lastRow; row++)
            {
                switch (row)
                {
                    case 2:
                        continue;
                    case 3:
                        var genWsColumn = directionColumns[directionColumnListIndex++];

                        var mergedRange = genWs.MergedCells[1, genWsColumn];
                        var address = new ExcelAddress(mergedRange);

                        totalDirections = address.Columns;

                        var dataColumnRange = totalDirections;
                        for (var column = genWsColumn; column <= genWsColumn + dataColumnRange - 1; column++)
                        {
                            var genData = genWs.Cells[row, column].Value?.ToString() ?? "";
                            if (genData.Contains("peds", StringComparison.InvariantCultureIgnoreCase)) totalDirections--;
                        }

                        var directionIndex = 0;
                        for (var column = 2; column < lastColumn; column++)
                        {
                            if (!DirectionTranslations.TryGetValue(allDirections[directionIndex++], out var translation))
                            {
                                HelperFunctions.WarningLog("Couldn't get direction translation");
                                translation = string.Empty;
                            }

                            toFormatWs.Cells[3, column].Value = translation == string.Empty ? allDirections[directionIndex - 1] : translation;
                        }

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

                            if (WorldDirectionTranslations.TryGetValue(data.EndsWith("bound") ? data.Remove(data.Length - 5, 5) : data, out var translation))
                            {
                                toFormatWs.Cells[2, column].Value = translation;

                                if (allDirections.Contains("peds cw", StringComparer.OrdinalIgnoreCase) || allDirections.Contains("peds ccw", StringComparer.OrdinalIgnoreCase))
                                {
                                    toFormatWs.Cells[1, column, 1, column + totalDirections + 1].Merge = true;
                                    toFormatWs.Cells[2, column, 2, column + totalDirections + 1].Merge = true;
                                }
                                else
                                {
                                    toFormatWs.Cells[1, column, 1, column + totalDirections - 1].Merge = true;
                                    toFormatWs.Cells[2, column, 2, column + totalDirections - 1].Merge = true;
                                }

                                column--;
                                continue;
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
        var lastRow = toFormatWs.Dimension.End.Row;
        var lastColumn = toFormatWs.Dimension.Columns;

        for (var row = 1; row <= lastRow; row++)
        {
            for (var column = 1; column <= lastColumn; column++)
            {
                var cellValue = toFormatWs.Cells[row, 1].Value?.ToString() ?? "";
                HelperFunctions.BorderAround(toFormatWs, row, column);

                if (string.IsNullOrWhiteSpace(cellValue) || !cellValue.Contains("%")) continue;

                var testValue = toFormatWs.Cells[row, column].Value?.ToString() ?? "";
                if (!decimal.TryParse(testValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var decimalValue)) continue;

                toFormatWs.Cells[row, column].Value = decimalValue;
                toFormatWs.Cells[row, column].Style.Numberformat.Format = "0.00%";
            }
        }

        var lastColumnRange = toFormatWs.Cells[1, lastColumn, lastRow, lastColumn].Style;
        lastColumnRange.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);

        lastColumnRange.Fill.PatternType = ExcelFillStyle.Solid;
        lastColumnRange.Fill.BackgroundColor.SetColor(Color.FromArgb(50,67, 255, 100));
        lastColumnRange.Fill.BackgroundColor.Tint = 0.5;

        toFormatWs.Cells[1, 1, 3, lastColumn].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);
        toFormatWs.Cells[1, 1, lastRow, 1].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);

        for (var row = 1; row <= lastRow; row++)
        {
            var cellValue = toFormatWs.Cells[row, 1].Value?.ToString() ?? "";
            if (!cellValue.Contains("spolu", StringComparison.OrdinalIgnoreCase)) continue;

            toFormatWs.Cells[row, 1, lastRow, lastColumn].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);
            toFormatWs.Cells[row, 1, row + 2, lastColumn].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);
            break;
        }

        for (var row = 1; row <= lastRow; row++)
        {
            var cellValue = toFormatWs.Cells[row, 1].Value?.ToString() ?? "";
            if (!cellValue.Contains("ch", StringComparison.OrdinalIgnoreCase) || cellValue.Contains("%")) continue;

            toFormatWs.Cells[row, 1, lastRow, lastColumn].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);
            break;
        }

        int columnsAmount;
        for (var column = 2; column <= lastColumn; column += columnsAmount)
        {
            var mergedRange = toFormatWs.MergedCells[1, column];
            columnsAmount = new ExcelAddress(mergedRange).Columns;

            toFormatWs.Cells[1, column, lastRow, column + columnsAmount - 1].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);
        }

        for (var column = 2; column <= lastColumn; column++)
        {
            var cellValue = toFormatWs.Cells[3, column].Value?.ToString() ?? "";
            if (!cellValue.Contains("peds", StringComparison.OrdinalIgnoreCase)) continue;
            for (var row = 3; row <= lastRow; row++)
            {
                toFormatWs.Cells[row, column].Style.Font.Color.SetColor(Color.Gray);
            }
        }

        toFormatWs.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        toFormatWs.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        toFormatWs.Cells.AutoFitColumns();
    }
}
