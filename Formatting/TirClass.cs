using System.Diagnostics;
using System.Drawing;
using ExcelFormatterConsole.Utility;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelFormatterConsole.Formatting;

public static class TirClass
// Total Intensity Rundown
{
    private const string ToFormatWorksheetName = "Celkový priebeh intenzít";

    private static readonly Stopwatch Stopwatch = Stopwatch.StartNew();

    private static readonly Dictionary<string, string> ViableVehicleCategoryTranslations = new(StringComparer.OrdinalIgnoreCase)
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

    private static readonly HashSet<string> ViableVehicleCategories = new(StringComparer.OrdinalIgnoreCase)
    {
        "motorcycles",
        "lights",
        "single-unit trucks",
        "articulated trucks",
        "buses",
        "bicycles on road",
        "articulated buses",
        "pedestrians",
    };

    private static readonly HashSet<string> AllKnownDirections = new(StringComparer.OrdinalIgnoreCase)
    {
        "north",
        "north-northeast",
        "northeast",
        "east-northeast",
        "east",
        "east-southeast",
        "southeast",
        "south-southeast",
        "south",
        "south-southwest",
        "southwest",
        "west-southwest",
        "west",
        "west-northwest",
        "northwest",
        "north-northwest"
    };


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

    private static void TimedLog(string logMessage)
    {
        Stopwatch.Stop();
        Console.WriteLine($"[{Stopwatch.Elapsed.TotalMilliseconds} ms] ----- {logMessage} -----");
        Stopwatch.Restart();
    }

    public static ExcelWorksheet FindCorrectWorksheet(ExcelPackage genPackage)
    {
        var eppWorksheet = genPackage.Workbook.Worksheets.FirstOrDefault(worksheet => worksheet.Index > 3 && worksheet.Name.Equals("total volume class breakdown", StringComparison.OrdinalIgnoreCase));
        return eppWorksheet ?? throw new MissingWorksheetException("TirClass.FindCorrectWorksheet() | Failed to find usable worksheet.");
    }

    public static ExcelWorksheet Prepare(ExcelPackage toFormatPackage)
    {
        var toFormatWs = toFormatPackage.Workbook.Worksheets.Add(ToFormatWorksheetName);
        TimedLog($" {toFormatWs.Name} | Created worksheet.");

        toFormatWs.Cells["A1:A3"].Merge = true;
        toFormatWs.Cells["A1"].Value = "Čas";

        return toFormatWs;
    }

    public static void FormatMeasuredTime(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {
        var row = 4;
        while (row < 1000)
        {
            var cellValue = genWs.Cells["A" + row].Value??"";
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
            var nextCellValue = genWs.Cells["A" + row].Value??"";

            if (string.IsNullOrWhiteSpace(nextCellValue.ToString()) || !DateTime.TryParse(nextCellValue.ToString(), out DateTime dtNextCellValue))
            {
                row--;
                var beforeLastRow = row - 1;

                cellValue = genWs.Cells["A" + beforeLastRow].Value;
                nextCellValue = genWs.Cells["A" + row].Value;

                var difference = DateTime.Parse(nextCellValue.ToString()??"00:00") - DateTime.Parse(cellValue.ToString()??"");

                nextCellValue = DateTime.Parse(nextCellValue.ToString()??"00:00").AddMinutes(difference.TotalMinutes);
                cellValue = DateTime.Parse(nextCellValue.ToString() ?? "00:00").AddMinutes(-difference.TotalMinutes);

                toFormatWs.Cells["A" + row].Value = $"{cellValue:HH:mm} - {nextCellValue:HH:mm}";

                break;
            }

            var correctCellRow = row - 1;
            toFormatWs.Cells["A" + correctCellRow].Value = $"{dtCellValue:HH:mm} - {dtNextCellValue:HH:mm}";
        }

        TimedLog($"{toFormatWs.Name} | Applied date formatting.");
    }


    public static int FormatVehicleCategories(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {
        List<string> vehicleCategoryTranslations = [];

        var lastRow = genWs.Dimension.End.Row;
        var firstRowAfterDates = toFormatWs.Dimension.End.Row;

        for (var row = firstRowAfterDates; row <= lastRow; row++)
        {
            var cellValue = genWs.Cells["A" + row].Value?.ToString()??"".ToLower().Trim();
            if (ViableVehicleCategoryTranslations.TryGetValue(cellValue, out var translation))
            {
                vehicleCategoryTranslations.Add(translation);
            }
        }

        for (var listIndex = 0; listIndex < vehicleCategoryTranslations.Count; listIndex++)
        {
            var column = listIndex + 2;
            var translation = vehicleCategoryTranslations[listIndex];

            toFormatWs.Cells[1, column].Value = translation;
            toFormatWs.Cells[1, column, 3, column].Merge = true;
        }

        TimedLog($"{toFormatWs.Name} | Applied vehicle category formatting.");
        return vehicleCategoryTranslations.Count;
    }

    public static Dictionary<string, Dictionary<string, double>> ReadPrimaryData(ExcelPackage genPackage, ExcelPackage toFormatPackage)
    {
        List<ExcelWorksheet> directionWorksheetsList =
            genPackage.Workbook.Worksheets.Where(ws => ws.Index > 0).Where(ws => AllKnownDirections.Contains(ws.Name.Replace("bound", "").Trim())).ToList();

        Dictionary<string, Dictionary<string, double>> primaryDataMapping = [];
        foreach (var ws in directionWorksheetsList)
        {
            var lastRow = ws.Dimension.End.Row;
            for (var row = 4; row <= lastRow; row++)
            {
                var lastColumn = ws.Dimension.End.Column;
                var cleanedDate = string.Empty;
                for (var column = 1; column <= lastColumn; column++)
                {
                    if (column == 1)
                    {
                        var dateKey = ws.Cells[row, column].Value?.ToString()??string.Empty;
                        if (!double.TryParse(dateKey, out var oaDate))
                        {
                            if (!DateTime.TryParse(dateKey, out var parsedDate))
                            {
                                throw new DateTimeConversionException("TirClass.ReadPrimaryData | Failed to convert into a DateTime object to be used as the first dictionary key.");
                            }

                            cleanedDate = $"{parsedDate:HH:mm}";
                        }
                        else
                        {
                            cleanedDate = $"{DateTime.FromOADate(oaDate):HH:mm}";
                        }

                        if (!primaryDataMapping.ContainsKey(cleanedDate))
                        {
                            primaryDataMapping[cleanedDate] = new Dictionary<string, double>();
                        }
                    }
                    else
                    {
                        const int categoriesRow = 3;
                        var category = ws.Cells[categoriesRow, column].Value?.ToString()??string.Empty;
                        if (string.IsNullOrWhiteSpace(category) || !ViableVehicleCategories.Contains(category))
                        {
                            throw new CategoryMatchException("TirClass.ReadPrimaryData | No valid category was found to be used as the second dictionary key.");
                        }

                        if (!VehicleCategoryTranslations.TryGetValue(category, out var translatedCategory))
                        {
                            throw new CategoryMatchException($"TirClass.ReadPrimaryData | No translation found for the category '{category}'.");
                        }

                        var cellValue = ws.Cells[row, column].Value?.ToString()??string.Empty;
                        if (string.IsNullOrWhiteSpace(cellValue))
                        {
                            continue;
                        }

                        if (!double.TryParse(cellValue, out var parsedCellValue))
                        {
                            throw new PrimaryDataValueException($"TirClass.ReadPrimaryData | Found an incorrect value: '{cellValue}' | row: {row} column: {column}.");
                        }

                        if (primaryDataMapping[cleanedDate].ContainsKey(translatedCategory))
                        {
                            primaryDataMapping[cleanedDate][translatedCategory] += parsedCellValue;
                        }
                        else
                        {
                            primaryDataMapping[cleanedDate][translatedCategory] = parsedCellValue;
                        }
                    }
                }
            }
        }

        TimedLog($"{genPackage.File.Name} | Read primary data from all directions.");
        return primaryDataMapping;
    }

    public static void WritePrimaryData(Dictionary<string, Dictionary<string, double>> primaryDataMapping, ExcelWorksheet toFormatWs)
    {
        const int categoryRow = 1;
        const int dateColumn = 1;

        var lastRow = toFormatWs.Dimension.End.Row;
        for (int row = 4; row <= lastRow; row++)
        {
            var date = (toFormatWs.Cells[row, dateColumn].Value?.ToString() ?? string.Empty).Split("-")[0].Trim();

            if (string.IsNullOrWhiteSpace(date) || date.Length > 5)
            {
                throw new UnexpectedValueException($"TirClass.WritePrimaryData | Encountered an unexpected value in the place of the date: '{date}'.");
            }

            var lastColumn = toFormatWs.Dimension.End.Column;
            for (var column = 2; column <= lastColumn; column++)
            {
                var category = toFormatWs.Cells[categoryRow, column].Value?.ToString()??string.Empty;

                if (string.IsNullOrWhiteSpace(category) || string.IsNullOrWhiteSpace(date))
                {
                    throw new UnexpectedValueException($"TirClass.WritePrimaryData | Encountered an unexpected value in the place of the category: '{category}'.");
                }

                toFormatWs.Cells[row, column].Value = primaryDataMapping[date][category];
            }
        }

        TimedLog($"{toFormatWs.Name} | Wrote all primary data.");
    }

    public static double CalculateAddedUpRowData(ExcelWorksheet toFormatWs)
    {
        var lastRow = toFormatWs.Dimension.End.Row + 2;
        var lastColumn = toFormatWs.Dimension.End.Column + 1;
        var rowToSkip = lastRow - 1;

        var everythingAddedUp = 0D;
        for (var column = 2; column <= lastColumn; column++)
        {
            if (column == lastColumn)
            {
                toFormatWs.Cells[lastRow, lastColumn].Value = everythingAddedUp;
                break;
            }

            var addedUpPrimaryData = 0D;
            for (var row = 4; row <= lastRow; row++)
            {
                if (row == rowToSkip)
                {
                    continue;
                }

                if (row == lastRow)
                {
                    toFormatWs.Cells[row, column].Value = addedUpPrimaryData;
                    break;
                }

                var cellValue = toFormatWs.Cells[row, column].Value?.ToString()??string.Empty;
                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    throw new UnexpectedValueException($"TirClass.CalculateAddedUpData | Detected an empty string in the cell | row: {row} column: {column}");
                }

                if (!double.TryParse(cellValue, out var parsedCellValue))
                {
                    throw new UnexpectedValueException($"TirClass.CalculateAddedUpData | Detected a non-numeric value in the cell | row: {row} column: {column}");
                }

                addedUpPrimaryData += parsedCellValue;
            }

            everythingAddedUp += addedUpPrimaryData;
        }

        TimedLog($"{toFormatWs.Name} | Calculated added up row data.");
        return everythingAddedUp;
    }

    public static double CalculateAddedUpColumnData(ExcelWorksheet toFormatWs)
    {
        var lastRow = toFormatWs.Dimension.End.Row - 1;
        var lastColumn = toFormatWs.Dimension.End.Column + 1;
        var columnToSkip = lastColumn - 1;

        var everythingAddedUp = 0D;
        for (var row = 4; row <= lastRow; row++)
        {
            if (row == lastRow)
            {
                toFormatWs.Cells[lastRow, lastColumn].Value = everythingAddedUp;
                break;
            }

            var addedUpPrimaryData = 0D;
            for (var column = 2; column <= lastColumn; column++)
            {
                if (column == columnToSkip)
                {
                    continue;
                }

                if (column == lastColumn)
                {
                    toFormatWs.Cells[row, column].Value = addedUpPrimaryData;
                    break;
                }

                var cellValue = toFormatWs.Cells[row, column].Value?.ToString()??string.Empty;
                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    throw new UnexpectedValueException($"TirClass.CalculateAddedUpData | Detected an empty string in the cell | row: {row} column: {column}");
                }

                if (!double.TryParse(cellValue, out var parsedCellValue))
                {
                    throw new UnexpectedValueException($"TirClass.CalculateAddedUpData | Detected a non-numeric value in the cell | row: {row} column: {column}");
                }

                addedUpPrimaryData += parsedCellValue;
            }

            everythingAddedUp += addedUpPrimaryData;
        }

        TimedLog($"{toFormatWs.Name} | Calculated added up column data.");
        return everythingAddedUp;
    }

    public static void CheckForMatchingResults(ExcelWorksheet toFormatWs, double addedUpRowData, double addedUpColumnData)
    {
        var lastRow = toFormatWs.Dimension.End.Row;
        var lastColumn = toFormatWs.Dimension.End.Column;

        var cellFill = toFormatWs.Cells[lastRow, lastColumn].Style.Fill;
        cellFill.PatternType = ExcelFillStyle.Solid;

        cellFill.BackgroundColor.SetColor((int)addedUpRowData == (int)addedUpColumnData ? Color.Green : Color.Red);
        TimedLog($"{toFormatWs.Name} | Performed a value check.");
    }


    /*
     public static void GenerateChart(ExcelPackage toFormatPackage)
    {
        using var stream = new MemoryStream();

        toFormatPackage.SaveAs(stream);
        stream.Position = 0;

        using XLWorkbook workbook = new XLWorkbook(stream);

        var ws = workbook.Worksheets.FirstOrDefault(worksheet => worksheet.Position > 3 && worksheet.Name.Equals("total volume class breakdown", StringComparison.OrdinalIgnoreCase));
        if (ws == null)
        {
            throw new MissingWorksheetException("TirClass.FindCorrectWorksheet() | Failed to find usable worksheet.");
        }

        var lastRow = ws.LastRowUsed().RowNumber() - 2;
        var lastColumn = ws.LastColumnUsed().ColumnNumber() - 2;

        var chart = ws
    }
    */


    public static void Styling(ExcelWorksheet toFormatWs)
    {
        toFormatWs.Cells.AutoFitColumns();
        toFormatWs.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        toFormatWs.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        toFormatWs.Cells["1:3"].Style.Font.Bold = true;
        toFormatWs.Cells["A:A"].Style.Font.Bold = true;

        var lastRow = toFormatWs.Dimension.End.Row;
        var lastColumn = toFormatWs.Dimension.End.Column - 1;

        for (var row = 1; row < lastRow; row++)
        {
            HelperFunctions.BorderAround(toFormatWs, row, lastColumn);
        }

        for (var column = 1; column < lastColumn; column++)
        {
            HelperFunctions.BorderAround(toFormatWs, lastRow, column);
        }

        lastColumn -= 2;
        lastRow -= 2;

        for (var column = 1; column <= lastColumn; column++)
        {
            for (var row = 1; row <= lastRow; row++)
            {
                HelperFunctions.BorderAround(toFormatWs, row, column);
            }
        }

        TimedLog($"{toFormatWs.Name} | Styled worksheet.");
    }
}
