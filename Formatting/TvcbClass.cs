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
        var genWs = genPackage.Workbook.Worksheets.FirstOrDefault(ws => ws.Index > 4 || ws.Name.ToLower() == "total volume class breakdown");
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

    public static void Navigation(ExcelWorksheet genWs, ExcelWorksheet toFormatWs, int directionsAmount)
    {
        toFormatWs.Cells["A1"].Value = "Smer od";
        toFormatWs.Cells["A2"].Value = "Orientácia";
        toFormatWs.Cells["A3"].Value = "Čas";

        var lastColumn = genWs.Dimension.End.Column;
        HashSet<string> allDirections = [];
        var fullDirectionAddresses = new Dictionary<string, string>();

        var row = 3;
        for (var column = 2; column <= lastColumn; column++)
        {
            var cellValue = genWs.Cells[row, column].Value?.ToString() ?? throw new UnexpectedValueException($"TvcbClass.Navigation | Unexpected null value detected | row={row} column={column}");
            if (!DirectionTranslations.TryGetValue(cellValue, out var directionTranslation))
            {
                HelperFunctions.ErrorLog($"TvcbClass.Navigation | Could not find valid translation for set direction | direction='{cellValue}'");
            }

            allDirections.Add(directionTranslation ?? cellValue);

            toFormatWs.Cells[row, column].Value = directionTranslation ?? cellValue;
        }

        row = 2;

        for (var column = 2; column <= lastColumn; column++)
        {
            toFormatWs.Cells[row, column].Value = "X - X";
        }

        row = 1;

        var specificRow = 3;
        var targetNumber = 1;

        while (targetNumber <=  directionsAmount)
        {
            for (var column = 2; column <= lastColumn; column++)
            {
                var cellValue = toFormatWs.Cells[specificRow, column].Value?.ToString() ?? throw new UnexpectedValueException($"TvcbClass.Navigation | Detected an unexpected null value | row={row} column={column}");

                if (cellValue.ToLower().Trim() != "doprava")
                {
                    continue;
                }

                cellValue = genWs.Cells[row, column].Value?.ToString() ?? throw new UnexpectedValueException($"TvcbClass.Navigation | Detected an unexpected null value | row={row} column={column}");

                var directionNumber = cellValue.Split("-")[0].Trim();
                if (!int.TryParse(directionNumber, out var verifiedNumber))
                {
                    continue;
                }

                var uncountedDirection = 1;
                var additionalOffset = 1;
                if (targetNumber == verifiedNumber)
                {
                    var fullAddress = ExcelCellBase.GetAddress(row, 1 + targetNumber * (allDirections.Count - uncountedDirection) - (allDirections.Count - uncountedDirection - additionalOffset));

                    toFormatWs.Cells[fullAddress].Value = cellValue;
                    targetNumber++;
                }

                try
                {
                    toFormatWs.Cells[row, column, row, column + allDirections.Count - 2].Merge = true;
                }
                catch (Exception)
                {
                    HelperFunctions.ErrorLog($"Failed to merge columns [{column} - {column + allDirections.Count}]");
                }
            }
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
