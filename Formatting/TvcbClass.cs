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
        var toFormatWs =toFormatPackage.Workbook.Worksheets.Add("Celkové údaje 12hod");
        toFormatWs.Cells.AutoFitColumns();

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

    public static void PrimaryDataReading(ExcelWorksheet genWs, int directions, int totalVehicleCategories)
    {
        var lastColumn = genWs.Dimension.End.Column;
        var lastRow = genWs.Dimension.End.Row + 1;

        directions++;

        List<string> lastColumnData = [];
        Dictionary<string, Dictionary<string, Dictionary<string, List<double>>>> primaryDataMapping = [];


        var mergedCellsColumn = 2;
        for (var column = 2; column <= lastColumn; column++)
        {
            string roadLegKey = string.Empty;
            string worldDirectionKey = string.Empty;
            string turnDirectionKey = string.Empty;
            List<double> primaryDataValues = [];


            for (var row = 1; row <= lastRow; row++)
            {
                const int rowOffset = 2;
                const int rowsPerCategory = 2;

                var specificEmptyRow = lastRow - (totalVehicleCategories * rowsPerCategory + rowOffset);

                if (column == lastColumn)
                {
                    if (row == specificEmptyRow || row == specificEmptyRow + 1 || row <= 3 || row == lastRow)
                    {
                        continue;
                    }

                    var lastColumnCellValue = genWs.Cells[row, column].Value?.ToString() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(lastColumnCellValue))
                    {
                        throw new UnexpectedValueException($"TvcbClass.PrimaryDataReading | Unexpected null value detected | Position: row={row} column={column} Location: name={genWs.Name} typeof={genWs.GetType()}");
                    }

                    lastColumnData.Add(lastColumnCellValue);
                    continue;
                }

                if (row == lastRow)
                {
                    if (string.IsNullOrWhiteSpace(roadLegKey) || string.IsNullOrWhiteSpace(worldDirectionKey) || string.IsNullOrWhiteSpace(turnDirectionKey))
                    {
                        throw new UnassignedVariableException($"TvcbClass.PrimaryDataReading | Unassigned keys detected before applying them to dictionary | Location: name={genWs.Name} typeof={genWs.GetType()}");
                    }

                    primaryDataMapping[roadLegKey] = new Dictionary<string, Dictionary<string, List<double>>>
                    {
                        [worldDirectionKey] = new()
                        {
                            [turnDirectionKey] = primaryDataValues
                        }
                    };

                    primaryDataValues = [];
                    continue;
                }

                string cellValue;
                if (row <= 3)
                {
                    cellValue = genWs.Cells[row, mergedCellsColumn].Value?.ToString() ?? string.Empty;
                }
                else
                {
                    cellValue = genWs.Cells[row, column].Value?.ToString() ?? string.Empty;
                }

                if (row == specificEmptyRow)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    throw new UnexpectedValueException($"TvcbClass.PrimaryDataReading | Unexpected null value detected | Position: row={row} column={column} Location: name={genWs.Name} typeof={genWs.GetType()}");
                }

                if (!double.TryParse(cellValue, out var parsedCellValue) && row <= 3)
                {
                    switch (row)
                    {
                        case 1:
                            roadLegKey = cellValue;
                            break;
                        case 2:
                            worldDirectionKey = cellValue;
                            break;
                        case 3:
                            turnDirectionKey = cellValue;
                            break;
                    }

                    continue;
                }

                if (!double.TryParse(cellValue, out _) && row > 3)
                {
                    throw new UnexpectedValueException($"TvcbClass.PrimaryDataReading | Unexpected non-numeric value detected | Position: row={row} column={column} Location: name={genWs.Name} typeof={genWs.GetType()}");
                }

                primaryDataValues.Add(parsedCellValue);
            }

            if (column % directions == 1)
            {
                mergedCellsColumn += directions;
            }
        }
    }
}