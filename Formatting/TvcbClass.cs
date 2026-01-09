using System.Diagnostics;
using ExcelFormatterConsole.Utility;
using OfficeOpenXml;

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

    public static ExcelWorksheet FindCorrectWorksheet(ExcelPackage genPackage)
    {
        var genWs = genPackage.Workbook.Worksheets.FirstOrDefault(ws => ws.Index > 4 || ws.Name.ToLower() == "total volume class breakdown");
        return genWs ?? throw new MissingWorksheetException($"TvcbClass.FindCorrectWorksheet | Failed to find correct worksheet. | Checked file name: '{genPackage.File.Name}'.");
    }

    public static ExcelWorksheet Prepare(ExcelPackage toFormatPackage)
    {
        var toFormatWs = toFormatPackage.Workbook.Worksheets.Add("Celkové údaje 12hod");
        toFormatWs.Cells.AutoFitColumns();

        return toFormatWs;
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

    public static void Navigation()
    {

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
}
