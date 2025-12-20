using System.Diagnostics;
using ExcelFormatterConsole.Utility;
using OfficeOpenXml;

namespace ExcelFormatterConsole.Formatting;

public static class TvcbClass
// Total Volume Class Breakdown
{
    private const string ToFormatWorksheetName = "Celkový priebeh intenzít";

    private static readonly Stopwatch Stopwatch = Stopwatch.StartNew();

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
        Console.WriteLine($"DEBUG: Looking at Worksheet: {genWs.Name}");
        Console.WriteLine($"DEBUG: Value in A4 is: '{genWs.Cells["A4"].Value}'");

        var row = 4;
        while (row < 1000)
        {
            // THE DATE IS NOT OADATE FORMAT

            var cellValue = genWs.Cells["A" + row].Value??"";
            if ( string.IsNullOrWhiteSpace(cellValue.ToString()) || ! double.TryParse(cellValue.ToString(), out var oaDate))
            {
                row++;
                continue;
            }

            cellValue = DateTime.FromOADate(oaDate);
            row++;

            var nextCellValue = genWs.Cells["A" + row].Value??"";

            if ( string.IsNullOrWhiteSpace(nextCellValue.ToString()) || ! double.TryParse(nextCellValue.ToString(), out var nextOaDate))
            {
                row--;
                var nextRow = row + 1;

                cellValue = genWs.Cells["A" + row].Value;
                nextCellValue = genWs.Cells["A" + nextRow].Value;

                var difference = DateTime.Parse(nextCellValue.ToString()) - DateTime.Parse(cellValue.ToString());

                nextCellValue = DateTime.Parse(cellValue.ToString()).AddMinutes(difference.TotalMinutes);
                toFormatWs.Cells["A" + row].Value = $"{cellValue:HH:mm} - {nextCellValue:HH:mm}";

                break;
            }

            nextCellValue = DateTime.FromOADate(nextOaDate);

            var correctCellRow = row - 1;
            toFormatWs.Cells["A" + correctCellRow].Value = $"{cellValue:HH:mm} - {nextCellValue:HH:mm}";
        }
    }

    /*
    public static void FormatVehicleCategories(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {

    }

    public static void ReadPrimaryData(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {

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

    public static void Styling(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {

    }
    */


}