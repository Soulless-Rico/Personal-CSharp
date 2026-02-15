using System.Drawing;
using ExcelFormatterConsole.Utility;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelFormatterConsole.Formatting;

public class DayPeaksClass
//  AM, PM Peaks
{
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

    private static ExcelWorksheet FindWorksheet(ExcelPackage genPackage, string worksheetName)
    {
        return genPackage.Workbook.Worksheets.First(ws => ws.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase));
    }
    public static void GenerateWorksheet (ExcelPackage toFormatPackage, ExcelPackage genPackage)
    {
        List<string> genWsNames = ["am peak class breakdown", "midday peak class breakdown", "pm peak class breakdown"];
        List<string> worksheetNames = ["Doobedňajšia hod.špička", "Obedňajšia hod.špička", "Poobedňajšia hod.špička"];

        Dictionary<string, string> secondaryNavigationCategories = new(StringComparer.OrdinalIgnoreCase)
        {
            ["grand total"] = "Spolu sk.voz.",
            ["% approach"] = "% pomer na smer",
            ["% total"] = "% celkový pomer",
            ["leg"] = "Smer od",
            ["direction"] = "Orientácia",
            ["start time"] = "Čas"
        };

        var index = 0;
        foreach (var wsName in worksheetNames)
        {
            var toFormatWs = toFormatPackage.Workbook.Worksheets.Add(wsName);
            var genWs = FindWorksheet(genPackage, genWsNames[index++]);

            List<string> mainDataColumn = [];
            List<string> secondaryDataColumn = [];
            List<string> totalDataColumn = [];
            List<int> mergedRanges = [];

            var lastRow = genWs.Dimension.End.Row;
            var lastColumn = genWs.Dimension.End.Column;

            var targetDirection = 1;

            for (var row = 1; row <= lastRow; row++)
            {
                var cellValue = genWs.Cells[row, 1].Value?.ToString() ?? "";
                secondaryDataColumn.Add(cellValue);

                if (row > 3) continue;

                restartColumn: ;

                for (var column = 2; column <= lastColumn; column++)
                {
                    if (column == lastColumn)
                    {
                        for (var innerRow = 1; innerRow <= lastRow; innerRow++)
                        {
                            cellValue = genWs.Cells[innerRow, lastColumn].Value?.ToString() ?? "";
                            totalDataColumn.Add(cellValue);
                        }
                        break;
                    }

                    cellValue = genWs.Cells[1, column].Value?.ToString() ?? "";

                    if (cellValue.Split("-")[0].Trim() != targetDirection.ToString()) continue;
                    targetDirection++;

                    var mergedRange = new ExcelAddress(genWs.MergedCells[row, column]).Columns;
                    mergedRanges.Add(mergedRange);

                    mergedRange = new ExcelAddress(genWs.MergedCells[row + 1, column]).Columns;
                    mergedRanges.Add(mergedRange);

                    for (var innerColumn = column; innerColumn <= column + mergedRange - 1; innerColumn++)
                    {
                        for (var innerRow = 1; innerRow <= lastRow; innerRow++)
                        {
                            cellValue = genWs.Cells[innerRow, innerColumn].Value?.ToString() ?? "";
                            mainDataColumn.Add(cellValue);
                        }
                    }

                    goto restartColumn;
                }
            }

            var listIndex = 0;
            var mergedIndex = 0;
            for (var column = 1; column <= lastColumn; column++)
            {
                for (var row = 1; row <= lastRow; row++)
                {
                    if (column == 1)
                    {
                        var data = secondaryDataColumn[listIndex++];
                        if (VehicleCategoryTranslations.TryGetValue(data, out var category))
                        {
                            data = category;
                        }
                        else if (VehicleCategoryTranslations.TryGetValue(data.Trim('%').Trim(), out var percentageCategory))
                        {
                            data = $"% {percentageCategory}";
                        }
                        else if (secondaryNavigationCategories.TryGetValue(data, out var navigation))
                        {
                            data = navigation;
                        }

                        if (data.Contains("phf", StringComparison.OrdinalIgnoreCase))
                        {
                            var timeStringList = data.Trim().Remove(0, 5).Trim(')').Split('-');
                            var leftTime = DateTime.TryParse(timeStringList[0], out var leftFormatedTime) ? $"{leftFormatedTime:HH:mm}" : timeStringList[0].Trim();
                            var rightTime = DateTime.TryParse(timeStringList[1], out var rightFormatedTime) ? $"{rightFormatedTime:HH:mm}" : timeStringList[1].Trim();

                            data = $"Špičk.hod ({leftTime} - {rightTime})";
                        }

                        toFormatWs.Cells[row, column].Value = data;
                        if (listIndex != secondaryDataColumn.Count) continue;
                    }
                    else if (column == lastColumn)
                    {
                        var data = totalDataColumn[listIndex++];
                        if (data.Contains("int total", StringComparison.OrdinalIgnoreCase))
                        {
                            data = "Celkom";
                        }

                        toFormatWs.Cells[row, lastColumn].Value = data;
                        if (listIndex != totalDataColumn.Count) continue;
                    }
                    else
                    {
                        var data = mainDataColumn[listIndex++];

                        if (row <= 3)
                        {
                            if (DirectionTranslations.TryGetValue(data, out var direction))
                            {
                                data = direction;
                            }
                            else if (data.Length > 5 && WorldDirectionTranslations.TryGetValue(data.Remove(data.Length - 5, 5), out var worldDirection))
                            {
                                data = worldDirection;
                            }
                        }

                        if (data != string.Empty && row is 1 or 2)
                        {
                            toFormatWs.Cells[row, column, row, column + mergedRanges[mergedIndex++] - 1].Merge = true;
                        }

                        toFormatWs.Cells[row, column].Value = data;
                        if (listIndex != mainDataColumn.Count) continue;
                    }

                    listIndex = 0;
                }
            }

            Style(toFormatWs);
        }
    }

    private static void Style(ExcelWorksheet toFormatWs)
    {
        var lastRow = toFormatWs.Dimension.End.Row;
        var lastColumn = toFormatWs.Dimension.End.Column;

        for (var row = 1; row <= lastRow; row++)
        {
            var cellValue = toFormatWs.Cells[row, 1].Value?.ToString() ?? "";
            if (!cellValue.Contains("%")) continue;

            for (var column = 1; column <= lastColumn; column++)
            {
                var testValue = toFormatWs.Cells[row, column].Value?.ToString() ?? "";
                if (!decimal.TryParse(testValue, out var decimalValue)) continue;

                toFormatWs.Cells[row, column].Value = decimalValue;
                toFormatWs.Cells[row, column].Style.Numberformat.Format = "0.00%";
            }
        }

        HelperFunctions.BorderEverythingInRange(toFormatWs, 1, 1, lastRow, lastColumn);

        const ExcelBorderStyle thickStyle = ExcelBorderStyle.Thick;
        var toFormatCells = toFormatWs.Cells;

        toFormatCells[1, 1, 3, lastColumn].Style.Border.BorderAround(thickStyle, Color.Black);
        toFormatCells[1, lastColumn, lastRow, lastColumn].Style.Border.BorderAround(thickStyle, Color.Black);
        toFormatCells[1, 1, lastRow, lastColumn].Style.Border.BorderAround(thickStyle, Color.Black);
        toFormatCells[1, 1, lastRow, 1].Style.Border.BorderAround(thickStyle, Color.Black);

        for (var row = 1; row <= lastRow; row++)
        {
            var cellValue = toFormatCells[row, 1].Value?.ToString() ?? "";
            if (cellValue.Contains("spolu", StringComparison.OrdinalIgnoreCase))
            {
                toFormatCells[4, 1, row, lastColumn].Style.Border.BorderAround(thickStyle, Color.Black);
            }

            if (!cellValue.Contains("ch", StringComparison.OrdinalIgnoreCase)) continue;
            toFormatCells[row, 1, lastRow, lastColumn].Style.Border.BorderAround(thickStyle, Color.Black);
            break;
        }

        for (var column = 2; column <= lastColumn; column++)
        {
            var cellValue = toFormatCells[2, column].Value?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(cellValue)) continue;

            var mergedColumns = new ExcelAddress(toFormatWs.MergedCells[2, column]).Columns;
            toFormatCells[1, column, lastRow, column + mergedColumns - 1].Style.Border.BorderAround(thickStyle, Color.Black);
        }

        for (var column = 2; column <= lastColumn; column++)
        {
            var cellValue = toFormatCells[3, column].Value?.ToString() ?? "";
            if (!cellValue.Contains("peds", StringComparison.OrdinalIgnoreCase)) continue;

            toFormatCells[3, column, lastRow, column].Style.Font.Color.SetColor(Color.Gray);
        }

        var fullLastColumn = toFormatCells[1, lastColumn, lastRow, lastColumn].Style.Fill;

        fullLastColumn.PatternType = ExcelFillStyle.Solid;
        fullLastColumn.BackgroundColor.SetColor(Color.FromArgb(50,67, 255, 100));
        fullLastColumn.BackgroundColor.Tint = 0.5;

        toFormatCells[1, lastColumn, 2, lastColumn].Merge = true;

        var style = toFormatWs.Cells.Style;
        style.VerticalAlignment = ExcelVerticalAlignment.Center;
        style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        toFormatWs.Cells.AutoFitColumns();
    }
}
