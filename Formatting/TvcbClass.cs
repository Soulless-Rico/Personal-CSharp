using System.Diagnostics;
using ExcelFormatterConsole.Utility;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelFormatterConsole.Formatting;

public static class TvcbClass
// Total Volume Class Breakdown
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


    private static void TimedLog(string logMessage)
    {
        Stopwatch.Stop();
        Console.WriteLine($"[{Stopwatch.Elapsed.TotalMilliseconds} ms] ----- {logMessage} -----");
        Stopwatch.Restart();
    }

    public static ExcelWorksheet FindCorrectWorksheet(ExcelPackage genPackage)
    {
        var passedWorksheet = genPackage.Workbook.Worksheets.FirstOrDefault(worksheet => worksheet.Index > 3 && worksheet.Name.Equals("total volume class breakdown", StringComparison.OrdinalIgnoreCase))!;
        return passedWorksheet == null! ?
            throw new MissingWorksheetException("TvcbClass.FindCorrectWorksheet() | Failed to find usable worksheet.") : passedWorksheet;
    }

    public static ExcelWorksheet Prepare(ExcelPackage toFormatPackage)
    {
        var toFormatWs = toFormatPackage.Workbook.Worksheets.Add(ToFormatWorksheetName);

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
    }


    public static void FormatVehicleCategories(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
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
    }

    public static void ReadPrimaryData(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {
        var lastRow = genWs.Dimension.End.Row;
        var firstRowBeforeDates = toFormatWs.Dimension.End.Row;

        for (var row = firstRowBeforeDates; row <= lastRow; row++)
        {
            var cellValue = genWs.Cells["A" + row].Value?.ToString() ?? "";
            if (!ViableVehicleCategories.Contains(cellValue))
            {
                continue;
            }

            List<string> primaryDataList = [];
            var lastColumn = genWs.Dimension.End.Column;
            for (var column = 1; column <= lastColumn; column++)
            {
                cellValue = genWs.Cells[row, column].Value?.ToString()??"0".Trim();
                if (ViableVehicleCategories.Contains(cellValue))
                {
                    continue;
                }

                primaryDataList.Add(cellValue);
            }
        }
    }

    public static void WritePrimaryData(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {

    }

    public static void CalculateAddedUpData(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {

    }

    public static void GenerateChart(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {

    }

    public static void Styling(ExcelWorksheet toFormatWs)
    {
        toFormatWs.Cells.AutoFitColumns();
        toFormatWs.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        toFormatWs.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        toFormatWs.Cells["1:3"].Style.Font.Bold = true;
        toFormatWs.Cells["A:A"].Style.Font.Bold = true;

    }



}
