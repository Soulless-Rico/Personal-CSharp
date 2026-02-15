using System.Drawing;
using System.Globalization;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ExcelFormatterConsole.Utility;

namespace ExcelFormatterConsole.Formatting;

public static class BasicDirectionsClass
{
    private const string FullDateFormat24Seconds = "dd-MM-yyyy HH:mm:ss";

    private static int _directions;

    public static List<string> AllDirections = [];

    private static readonly Dictionary<string, string> CellMapping = new ()
    {
        { "A1", "Názov štúdie" }, { "B1", "B1" },
        { "A2", "Projekt" }, { "B2", "B2" },
        { "A3", "Kód projektu" }, { "B3", "B3" },
        { "A4", "Smery a vozidlá" }, { "B4", "B4" },
        { "A5", "Časové intervaly" }, { "B5", "B5" },
        { "A6", "Časová zóna" }, { "B6", "B6" },

        { "A7", "Začiatok merania" },
        { "A8", "Koniec merania" },

        { "A9", "Miesto" }, { "B9", "B9" },
        { "A10", "Latitude and Longitude (GPS  LAT / LON)" }, { "B10", "B10" },
        { "A12", "Doobed. špička" }, { "B12", "B12" },
        { "A13", "Stredná špička" }, { "B13", "B13" },
        { "A14", "Poobedná špička" }, { "B14", "B14" },
        { "A16", "Poznámka 1" }, { "B16", "B16" },
        { "A17", "Poznámka 2" }, { "B17", "B17" },
        { "A18", "Poznámka 3" }, { "B18", "B18" },
        { "A19", "Poznámka 4" }, { "B19", "B20" },
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
    };

    private static readonly Dictionary<string, string> VehicleCategories = new(StringComparer.OrdinalIgnoreCase)
    {
        {"motorcycles", "M"},
        {"lights", "LV"},
        {"single-unit trucks", "NV"},
        {"articulated trucks", "TNV"},
        {"buses", "A"},
        {"bicycles on road", "B"},
        {"articulated buses", "AK"},
        {"pedestrians", "CH"},
        {"spolu", "Spolu"}
    };

    private static readonly HashSet<string> AllKnownDirections =
    [
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
    ];

    private static (List<string>, int) DetermineVehicleCategories(ExcelWorksheet worksheetObject)
    {
        var lastVehicleCategory = -1;

        List<string> vehicleCategories = [];
        for (var letterIndex = 2;; letterIndex++)
        {
            var currentCellValue = worksheetObject.Cells[3, letterIndex].Value?.ToString() ?? string.Empty;
            if (vehicleCategories.Contains(currentCellValue))
            {
                vehicleCategories.Add("Spolu");
                lastVehicleCategory++;
                break;
            }

            lastVehicleCategory++;
            vehicleCategories.Add(currentCellValue);
        }

        List<string> categoryShortcuts = [];
        foreach (var category in vehicleCategories)
        {
            if (VehicleCategories.TryGetValue(category, out var shortcut))
            {
                categoryShortcuts.Add(shortcut);
            }
            else
            {
                categoryShortcuts.Add(category);
                HelperFunctions.ErrorLog($"Category Error: no match found for {category}");
            }
        }

        return (categoryShortcuts, lastVehicleCategory);
    }

    public static void DefaultData(ExcelWorksheet genWs, ExcelWorksheet toFormatWs)
    {
        toFormatWs.Name = "Základné údaje";

        foreach (var mapping in CellMapping)
        {
            var cell = mapping.Key;
            var value = mapping.Value;

            if (cell.StartsWith("A"))
            {
                toFormatWs.Cells[cell].Value = value;
            }

            if (cell.StartsWith("B"))
            {
                toFormatWs.Cells[cell].Value = genWs.Cells[cell].Value;
            }
        }

        try
        {
            var dt1 = DateTime.FromOADate(Convert.ToDouble(genWs.Cells["B7"].Value));
            var dt2 = DateTime.FromOADate(Convert.ToDouble(genWs.Cells["B8"].Value));

            toFormatWs.Cells["B7"].Value =
                dt1.ToString(FullDateFormat24Seconds, CultureInfo.InvariantCulture);
            toFormatWs.Cells["B8"].Value =
                dt2.ToString(FullDateFormat24Seconds, CultureInfo.InvariantCulture);
        }
        catch (FormatException exception)
        {
            Console.WriteLine(exception);
            throw;
        }

        toFormatWs.Cells[toFormatWs.Dimension.Address].AutoFitColumns();
        toFormatWs.Cells["B7:B8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
    }

    public static int FindAllDirections(ExcelPackage genPackage)
    {
        var directionSheets = genPackage.Workbook.Worksheets.Where(ws => ws.Index > 0).Where(ws =>
        {
            var cleanedName = ws.Name.Trim().ToLower();

            if (cleanedName.EndsWith("bound"))
            {
                cleanedName = cleanedName.Remove(cleanedName.Length - 5, 5).Trim();
            }

            return AllKnownDirections.Contains(cleanedName);
        }).ToList();

        _directions = directionSheets.Count;

        for (var cycle = 1; cycle <= _directions; cycle++)
        {

            var selectedWorksheet = genPackage.Workbook.Worksheets[cycle];
            var cellValue = selectedWorksheet.Cells["B1"].Value?.ToString() ?? string.Empty.Trim();
            var directionCode = cellValue.Split("-");

            cellValue = directionCode[0].Trim();
            switch (cellValue)
            {
                case "1" or "2" or "3" or "4" or "5" or "6":
                    AllDirections.Add(cellValue + ":" + selectedWorksheet.Index);
                    break;
                default:
                    throw new MissingDirectionMach(
                        $"Unrecognized direction code found: {cellValue} in worksheet {selectedWorksheet.Name}");
            }
        }

        return _directions;
    }

    public static void BasicDirections(int directionNumber, ExcelPackage genPackage, ExcelPackage toFormatPackage)
    {
        var toFormatWs = toFormatPackage.Workbook.Worksheets.Add(directionNumber.ToString());
        var genWs = (from direction in AllDirections
            where direction.Contains(directionNumber + ":")
            select genPackage.Workbook.Worksheets[Convert.ToInt16(direction.Trim().Remove(0, 2))]).FirstOrDefault(ws => ws != null);

        if (genWs == null)
        {
            HelperFunctions.ErrorLog($"Couldn't find correct worksheet for direction: {directionNumber}");
            return;
        }

        toFormatWs.Cells["B1"].Value = genWs.Cells["B1"].Value;
        toFormatWs.Name = genWs.Cells["B1"].Value.ToString();
        toFormatWs.Cells["A1"].Value = "Čas";

        var lastRow = genWs.Dimension.End.Row;
        var inputRow = 4;
        while (true)
        {
            if (inputRow == lastRow)
            {
                inputRow--;
                var lastCellTimeValue = toFormatWs.Cells["A" + inputRow].Value?.ToString() ?? "00:00-00:00";
                var lastCellTime = lastCellTimeValue.Split("-");

                if (lastCellTime.Length < 2)
                {
                    HelperFunctions.ErrorLog("Time Formatting Error: Could not split last time interval");
                    break;
                }

                if (!DateTime.TryParse(lastCellTime[1].Trim(), out var t2))
                {
                    HelperFunctions.ErrorLog("DateTime TryParse Failure");
                    break;
                }

                var difference = DateTime.Parse(lastCellTime[1].Trim()) -
                                 DateTime.Parse(lastCellTime[0].Trim());

                inputRow++;
                toFormatWs.Cells["A" + inputRow].Value =
                    $"{lastCellTime[1].Trim()} - {t2.AddMinutes(difference.TotalMinutes):HH:mm}";

                break;
            }

            var currentCell = genWs.Cells["A" + inputRow].Value;
            var nextCell = genWs.Cells["A" + (inputRow + 1)].Value;

            if (currentCell == null || nextCell == null)
            {
                break;
            }

            var dt1 = DateTime.FromOADate(Convert.ToDouble(currentCell));
            var dt2 = DateTime.FromOADate(Convert.ToDouble(nextCell));

            toFormatWs.Cells["A" + inputRow].Value =
                $"{dt1.ToString("HH:mm", CultureInfo.InvariantCulture)} - {dt2.ToString("HH:mm", CultureInfo.InvariantCulture)}";

            inputRow++;
        }

        var lastRowBeforeExpandingNewWorksheet = toFormatWs.Dimension.End.Row;

        List<string> primaryData = [];
        for (var letterIndex = 2; letterIndex <= genWs.Dimension.End.Column; letterIndex++)
        {
            for (var rowIndex = 4; rowIndex <= genWs.Dimension.End.Row; rowIndex++)
            {
                primaryData.Add(genWs.Cells[rowIndex, letterIndex].Value.ToString() ?? string.Empty);

                if (genWs.Cells[rowIndex, letterIndex].Value.ToString() == string.Empty)
                {
                    HelperFunctions.ErrorLog($"Null Value Detected: Primary data writing detected an empty value| row:{rowIndex} column:{letterIndex} |");
                }
            }
        }

        var (categoryShortcuts, lastVehicleCategory) = DetermineVehicleCategories(genWs);
        var totalCategories = lastVehicleCategory + 1;

        var listIndex = 0;
        var leftOutColumns = 1;
        for (var letterIndex = 2;
             letterIndex < genWs.Dimension.End.Column + leftOutColumns;
             letterIndex++)
        {
            if ((letterIndex - 1) % totalCategories == 0)
            {
                leftOutColumns++;
                continue;
            }

            for (var rowIndex = 4; rowIndex <= genWs.Dimension.End.Row; rowIndex++)
            {
                toFormatWs.Cells[rowIndex, letterIndex].Value = primaryData[listIndex];
                listIndex++;
            }
        }

        listIndex = 0;
        var lastCategoryCell = toFormatWs.Dimension.End.Column + 1;
        for (var letterIndex = 2; letterIndex <= lastCategoryCell; letterIndex++)
        {
            if (listIndex <= lastVehicleCategory)
            {
                toFormatWs.Cells[3, letterIndex].Value = categoryShortcuts[listIndex];
                listIndex++;
            }
            else
            {
                letterIndex--;
                listIndex = 0;
            }
        }

        for (var rowIndex = 4; rowIndex <= lastRowBeforeExpandingNewWorksheet; rowIndex++)
        {
            var addedTogether = 0;
            for (var letterIndex = 2; letterIndex <= toFormatWs.Dimension.End.Column; letterIndex++)
            {
                if ((letterIndex - 1) % totalCategories == 0)
                {
                    toFormatWs.Cells[rowIndex, letterIndex].Value = addedTogether;

                    addedTogether = 0;
                    continue;
                }

                var cellValue = toFormatWs.Cells[rowIndex, letterIndex].Value?.ToString() ?? "0".Trim('`').Trim();
                if (int.TryParse(cellValue, out var number))
                {
                    addedTogether += number;
                }
                else
                {
                    HelperFunctions.ErrorLog($"Non-Numeric Value Detected | row:{rowIndex} column:{letterIndex} |");
                }
            }
        }

        var gapBetweenWorksheets = 0;
        for (var columnIndex = 2; columnIndex <= toFormatWs.Dimension.End.Column; columnIndex++)
        {
            if ((columnIndex - 1) % totalCategories == 0)
            {
                toFormatWs.Cells[2, columnIndex - lastVehicleCategory, 2, columnIndex].Merge = true;
            }
            else if (columnIndex == 2)
            {
                toFormatWs.Cells[2, 2, 2, totalCategories + 1].Merge = true;

                if (string.IsNullOrWhiteSpace(genWs.Cells[2, columnIndex - gapBetweenWorksheets].Value.ToString()))
                {
                    HelperFunctions.ErrorLog("direction is null, skipping formatting");
                    continue;
                }

                var key = genWs.Cells[2, columnIndex - gapBetweenWorksheets].Value?.ToString() ?? "0 - 0 error".ToLower();
                toFormatWs.Cells[2, columnIndex].Value = DirectionTranslations.TryGetValue(key, out var translation) ?
                    translation : genWs.Cells[2, columnIndex - gapBetweenWorksheets].Value;

                if (key == "u-turn")
                {
                    toFormatWs.Cells[2, columnIndex].Value = translation;
                }

                gapBetweenWorksheets++;
            }

            if ((columnIndex - 2) % totalCategories != 0 || columnIndex == 2) continue;
            {
                if (string.IsNullOrWhiteSpace(genWs.Cells[2, columnIndex - gapBetweenWorksheets].Value.ToString()))
                {
                    HelperFunctions.ErrorLog("direction is null, skipping formatting");
                    continue;
                }

                var key = genWs.Cells[2, columnIndex - gapBetweenWorksheets].Value?.ToString() ?? "0 - 0 error".ToLower();
                toFormatWs.Cells[2, columnIndex].Value = DirectionTranslations.TryGetValue(key, out var translation) ?
                    translation : genWs.Cells[2, columnIndex - gapBetweenWorksheets].Value;

                if (key == "u-turn")
                {
                    toFormatWs.Cells[2, columnIndex].Value = translation;
                }

                gapBetweenWorksheets++;
            }
        }

        for (var letterIndex = 2; letterIndex <= toFormatWs.Dimension.End.Column; letterIndex++)
        {
            if ((letterIndex - 1) % totalCategories != 0) continue;
            var addedTogether = 0;
            for (var rowIndex = 4; rowIndex <= lastRowBeforeExpandingNewWorksheet + 1; rowIndex++)
            {
                if (rowIndex == lastRowBeforeExpandingNewWorksheet + 1)
                {
                    toFormatWs.Cells[rowIndex + 1, letterIndex].Value = addedTogether;
                    break;
                }

                addedTogether += Convert.ToInt32(toFormatWs.Cells[rowIndex, letterIndex].Value.ToString() ?? "0".Trim('`').Trim());
            }
        }

        var lastColumn = toFormatWs.Dimension.End.Column + 1;
        toFormatWs.Cells[1, lastColumn, 3, lastColumn].Merge = true;
        toFormatWs.Cells[1, lastColumn, 3, lastColumn].Value = "Suma";

        var onlySelectedColumns = false;
        for (var rowIndex = 4; rowIndex <= lastRowBeforeExpandingNewWorksheet + 2; rowIndex++)
        {
            if (rowIndex == lastRowBeforeExpandingNewWorksheet + 1)
            {
                continue;
            }

            if (rowIndex == lastRowBeforeExpandingNewWorksheet + 2)
            {
                onlySelectedColumns = true;
            }

            var addedTogether = 0;
            for (var letterIndex = 2; letterIndex <= lastColumn; letterIndex++)
            {
                if (letterIndex == lastColumn)
                {
                    toFormatWs.Cells[rowIndex, letterIndex].Value = addedTogether;
                    break;
                }

                if ((letterIndex - 1) % totalCategories != 0 && !onlySelectedColumns || (letterIndex - 1) % totalCategories == 0 && onlySelectedColumns)
                {
                    addedTogether += Convert.ToInt32(toFormatWs.Cells[rowIndex, letterIndex].Value.ToString() ?? "0".Trim('`').Trim());
                }
            }
        }

        toFormatWs.Cells[1, 1, 3, toFormatWs.Dimension.End.Column].Style.Font.Bold = true;
        toFormatWs.Cells[1, 1, 3, toFormatWs.Dimension.End.Column].Style.HorizontalAlignment =
            ExcelHorizontalAlignment.Center;

        toFormatWs.Cells[1, 2, 1, toFormatWs.Dimension.End.Column - 1].Merge = true;

        toFormatWs.Cells["A1:A3"].Merge = true;
        toFormatWs.Cells["A1:A3"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);
        toFormatWs.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        toFormatWs.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


        for (var columnIndex = 2; columnIndex <= toFormatWs.Dimension.End.Column; columnIndex++)
        {
            if ((columnIndex - 1) % totalCategories != 0) continue;
            toFormatWs.Cells[2, columnIndex - lastVehicleCategory, 2, columnIndex].Style.Border
                .BorderAround(ExcelBorderStyle.Thick, Color.Black);

            toFormatWs.Cells[3, columnIndex - lastVehicleCategory, 3, columnIndex].Style.Border
                .BorderAround(ExcelBorderStyle.Thick, Color.Black);
        }

        toFormatWs.Cells[1, toFormatWs.Dimension.End.Column, 3, toFormatWs.Dimension.End.Column].Style.Border
            .BorderAround(ExcelBorderStyle.Thick);

        toFormatWs.Cells["A1:A" + lastRowBeforeExpandingNewWorksheet].Style.Border
            .BorderAround(ExcelBorderStyle.Thick);

        for (var row = 4; row < lastRowBeforeExpandingNewWorksheet; row++)
        {
            var cellValue = toFormatWs.Cells["A" + row].Value?.ToString() ?? "00:00-00:00";
            var cellValueSplit = cellValue.Split("-");

            var dt = DateTime.Parse(cellValueSplit[1]);
            if (dt.ToString("mm") == "00")
            {
                toFormatWs.Cells[row, 1, row, toFormatWs.Dimension.End.Column - 1].Style.Border.Bottom.Style =
                    ExcelBorderStyle.Thick;
            }
        }

        toFormatWs.Cells[1, 1, lastRowBeforeExpandingNewWorksheet, toFormatWs.Dimension.End.Column - 1].Style
            .Border.BorderAround(ExcelBorderStyle.Thick);

        for (var columnIndex = 2; columnIndex <= toFormatWs.Dimension.End.Column; columnIndex++)
        {
            if ((columnIndex - 1) % totalCategories == 0)
            {
                toFormatWs.Cells[3, columnIndex, lastRowBeforeExpandingNewWorksheet, columnIndex].Style.Border
                    .BorderAround(ExcelBorderStyle.Thick, Color.Black);
            }
        }

        var greyColorRows = lastRowBeforeExpandingNewWorksheet + 2;
        for (var columnIndex = 2; columnIndex <= toFormatWs.Dimension.End.Column; columnIndex++)
        {
            if ((columnIndex - 1) % totalCategories != 0) continue;
            toFormatWs.Cells[3, columnIndex, greyColorRows, columnIndex].Style.Fill.PatternType =
                ExcelFillStyle.Solid;
            toFormatWs.Cells[3, columnIndex, greyColorRows, columnIndex].Style.Fill.BackgroundColor
                .SetColor(Color.LightGray);
        }

        var greenColorColumns = toFormatWs.Dimension.End.Column - 1;
        for (var rowIndex = 1; rowIndex <= toFormatWs.Dimension.End.Row; rowIndex++)
        {
            toFormatWs.Cells[rowIndex, greenColorColumns].Style.Fill.PatternType = ExcelFillStyle.Solid;
            toFormatWs.Cells[rowIndex, greenColorColumns].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(50,67, 255, 100));
            toFormatWs.Cells[rowIndex, greenColorColumns].Style.Fill.BackgroundColor.Tint = 0.5;
        }

        toFormatWs.Cells["A:A"].Style.Font.Bold = true;

        toFormatWs.Cells["1:3"].Style.Font.Bold = true;

        toFormatWs.Cells["A1:" + toFormatWs.Dimension.End.Address].Style.Numberformat.Format = "General";
        toFormatWs.Cells["A1:" + toFormatWs.Dimension.End.Address].AutoFitColumns();
        toFormatWs.Cells["A1:" + toFormatWs.Dimension.End.Address].Style.HorizontalAlignment =
            ExcelHorizontalAlignment.Center;
        toFormatWs.Cells["A1:" + toFormatWs.Dimension.End.Address].Style.VerticalAlignment =
            ExcelVerticalAlignment.Center;
    }
}
