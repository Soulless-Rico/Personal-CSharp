using System.Diagnostics;
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
        { "right", "smer doprava" },
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

    private static readonly Stopwatch Stopwatch = Stopwatch.StartNew();

    public static void TimedLog(string logMessage)
    {
        Stopwatch.Stop();
        Console.WriteLine($"[{Stopwatch.Elapsed.TotalMilliseconds} ms] ----- {logMessage} -----");
        Stopwatch.Restart();
    }

    private static (List<string>, int) DetermineVehicleCategories(ExcelWorksheet worksheetObject)
    {
        var lastVehicleCategory = -1;

        // getting vehicle categories
        List<string> vehicleCategories = [];
        for (var letterIndex = 2;; letterIndex++)
        {
            var currentCellValue = worksheetObject.Cells[3, letterIndex].Value?.ToString() ?? string.Empty;
            if (vehicleCategories.Contains(currentCellValue))
            {
                // check so no duplicates get added + added additional category
                vehicleCategories.Add("Spolu");
                lastVehicleCategory++;
                break;
            }

            lastVehicleCategory++;
            vehicleCategories.Add(currentCellValue);
        }

        TimedLog($"{worksheetObject.Name} | Getting vehicle categories 50%");

        // applying shortcuts
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

        TimedLog($"{worksheetObject.Name} | Applied vehicle categories 100%");
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

        TimedLog($"{toFormatWs.Name} | Cell mapping 25%");

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

        TimedLog($"{toFormatWs.Name} | Date formatting 50%");

        toFormatWs.Cells[toFormatWs.Dimension.Address].AutoFitColumns();
        toFormatWs.Cells["B7:B8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

        TimedLog($"{toFormatWs.Name} | Styling worksheet 100%");
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

        TimedLog($"{genPackage.File.Name} | Finding all directions 50%");

        _directions = directionSheets.Count;

        // loop through all found directions
        for (var cycle = 1; cycle <= _directions; cycle++)
        {
            // read worksheet data to determine correct direction

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

        TimedLog($"{genPackage.File.Name} | Assigning direction values 100%");
        return _directions;
    }

    public static void BasicDirections(int directionNumber, ExcelPackage genPackage, ExcelPackage toFormatPackage)
    {
        var newWorksheet = toFormatPackage.Workbook.Worksheets.Add(directionNumber.ToString());
        ExcelWorksheet? generatedWorksheet = null;

        // find correct worksheet -------------------------------------------------------------------------------------------------------------------------------------------------

        foreach (var direction in AllDirections)
        {
            // check for correct direction
            if (direction.Contains(directionNumber + ":"))
            {
                var sheet = genPackage.Workbook.Worksheets[Convert.ToInt16(direction.Trim().Remove(0, 2))];
                if (sheet != null)
                {
                    generatedWorksheet = sheet;
                    break;
                }
            }
        }

        if (generatedWorksheet == null)
        {
            HelperFunctions.ErrorLog($"Couldn't find correct worksheet for direction: {directionNumber}");
            return;
        }

        newWorksheet.Cells["B1"].Value = generatedWorksheet.Cells["B1"].Value;
        newWorksheet.Name = generatedWorksheet.Cells["B1"].Value.ToString();
        newWorksheet.Cells["A1"].Value = "Čas";

        TimedLog($"{newWorksheet.Name} | Finding correct worksheet 20%");

        // time formatting -------------------------------------------------------------------------------------------------------------------------------------------------

        var lastRow = generatedWorksheet.Dimension.End.Row;
        var inputRow = 4;
        while (true)
        {
            if (inputRow == lastRow)
            {
                inputRow--;
                var lastCellTimeValue = newWorksheet.Cells["A" + inputRow].Value?.ToString() ?? "00:00-00:00";
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
                newWorksheet.Cells["A" + inputRow].Value =
                    $"{lastCellTime[1].Trim()} - {t2.AddMinutes(difference.TotalMinutes):HH:mm}";

                //Console.WriteLine($"Loop broken on last row: {inputRow}");
                break;
            }

            var currentCell = generatedWorksheet.Cells["A" + inputRow].Value;
            var nextCell = generatedWorksheet.Cells["A" + (inputRow + 1)].Value;

            if (currentCell == null || nextCell == null)
            {
                //Console.WriteLine($"Loop broken on row: {inputRow}");
                break;
            }

            var dt1 = DateTime.FromOADate(Convert.ToDouble(currentCell));
            var dt2 = DateTime.FromOADate(Convert.ToDouble(nextCell));

            newWorksheet.Cells["A" + inputRow].Value =
                $"{dt1.ToString("HH:mm", CultureInfo.InvariantCulture)} - {dt2.ToString("HH:mm", CultureInfo.InvariantCulture)}";

            inputRow++;
        }

        var lastRowBeforeExpandingNewWorksheet = newWorksheet.Dimension.End.Row;

        TimedLog($"{newWorksheet.Name} | Formatted special cells 40%");

        // primary data part 1 -------------------------------------------------------------------------------------------------------------------------------------------------

        // reading primary data
        List<string> primaryData = [];
        for (var letterIndex = 2; letterIndex <= generatedWorksheet.Dimension.End.Column; letterIndex++)
        {
            for (var rowIndex = 4; rowIndex <= generatedWorksheet.Dimension.End.Row; rowIndex++)
            {
                primaryData.Add(generatedWorksheet.Cells[rowIndex, letterIndex].Value.ToString() ?? string.Empty);

                if (generatedWorksheet.Cells[rowIndex, letterIndex].Value.ToString() == string.Empty)
                {
                    HelperFunctions.ErrorLog($"Null Value Detected: Primary data writing detected an empty value| row:{rowIndex} column:{letterIndex} |");
                }
            }
        }

        // Other
        var (categoryShortcuts, lastVehicleCategory) = DetermineVehicleCategories(generatedWorksheet);
        var totalCategories = lastVehicleCategory + 1;

        // primary data part 2 ------------------------------------------------------------------------------------------------------------------------------------

        // writing primary data
        var listIndex = 0;
        var leftOutColumns = 1;
        for (var letterIndex = 2;
             letterIndex < generatedWorksheet.Dimension.End.Column + leftOutColumns;
             letterIndex++)
        {
            if ((letterIndex - 1) % totalCategories == 0)
            {
                leftOutColumns++;
                continue;
            }

            for (var rowIndex = 4; rowIndex <= generatedWorksheet.Dimension.End.Row; rowIndex++)
            {
                newWorksheet.Cells[rowIndex, letterIndex].Value = primaryData[listIndex];
                listIndex++;
            }
        }

        TimedLog($"{newWorksheet.Name} | Read and wrote all primary data 60%");

        // writing vehicle categories in correct spots
        listIndex = 0;
        var lastCategoryCell = newWorksheet.Dimension.End.Column + 1;
        for (var letterIndex = 2; letterIndex <= lastCategoryCell; letterIndex++)
        {
            if (listIndex <= lastVehicleCategory)
            {
                newWorksheet.Cells[3, letterIndex].Value = categoryShortcuts[listIndex];
                listIndex++;
            }
            else
            {
                letterIndex--;
                listIndex = 0;
            }
        }

        // added up primary data
        for (var rowIndex = 4; rowIndex <= lastRowBeforeExpandingNewWorksheet; rowIndex++)
        {
            var addedTogether = 0;
            for (var letterIndex = 2; letterIndex <= newWorksheet.Dimension.End.Column; letterIndex++)
            {
                if ((letterIndex - 1) % totalCategories == 0)
                {
                    newWorksheet.Cells[rowIndex, letterIndex].Value = addedTogether;

                    addedTogether = 0;
                    continue;
                }

                var cellValue = newWorksheet.Cells[rowIndex, letterIndex].Value?.ToString() ?? "0".Trim('`').Trim();
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

        // 2nd row part -------------------------------------------------------------------------------------------------------------------------------------------------

        var gapBetweenWorksheets = 0;
        for (var columnIndex = 2; columnIndex <= newWorksheet.Dimension.End.Column; columnIndex++)
        {
            if ((columnIndex - 1) % totalCategories == 0)
            {
                newWorksheet.Cells[2, columnIndex - lastVehicleCategory, 2, columnIndex].Merge = true;
            }
            else if (columnIndex == 2)
            {
                newWorksheet.Cells[2, 2, 2, totalCategories + 1].Merge = true;

                if (string.IsNullOrWhiteSpace(generatedWorksheet.Cells[2, columnIndex - gapBetweenWorksheets].Value.ToString()))
                {
                    HelperFunctions.ErrorLog("direction is null, skipping formatting");
                    continue;
                }

                var key = generatedWorksheet.Cells[2, columnIndex - gapBetweenWorksheets].Value?.ToString() ?? "0 - 0 error".ToLower();
                newWorksheet.Cells[2, columnIndex].Value = DirectionTranslations.TryGetValue(key, out var translation) ?
                    $"{directionNumber} - {gapBetweenWorksheets + 1} {translation}" : generatedWorksheet.Cells[2, columnIndex - gapBetweenWorksheets].Value;

                if (key == "u-turn")
                {
                    newWorksheet.Cells[2, columnIndex].Value = $"{directionNumber} - {directionNumber} {translation}";
                }

                gapBetweenWorksheets++;
            }

            if ((columnIndex - 2) % totalCategories == 0 && columnIndex != 2)
            {
                if (string.IsNullOrWhiteSpace(generatedWorksheet.Cells[2, columnIndex - gapBetweenWorksheets].Value.ToString()))
                {
                    HelperFunctions.ErrorLog("direction is null, skipping formatting");
                    continue;
                }

                var key = generatedWorksheet.Cells[2, columnIndex - gapBetweenWorksheets].Value?.ToString() ?? "0 - 0 error".ToLower();
                newWorksheet.Cells[2, columnIndex].Value = DirectionTranslations.TryGetValue(key, out var translation) ?
                    $"{directionNumber} - {gapBetweenWorksheets + 1} {translation}" : generatedWorksheet.Cells[2, columnIndex - gapBetweenWorksheets].Value;

                if (key == "u-turn")
                {
                    newWorksheet.Cells[2, columnIndex].Value = $"{directionNumber} - {directionNumber} {translation}";
                }

                //newWorksheet.Cells[2, columnIndex].Value = generatedWorksheet.Cells[2, columnIndex - gapBetweenWorksheets].Value;
                gapBetweenWorksheets++;
            }
        }

        // added up column data
        for (var letterIndex = 2; letterIndex <= newWorksheet.Dimension.End.Column; letterIndex++)
        {
            if ((letterIndex - 1) % totalCategories == 0)
            {
                var addedTogether = 0;
                for (var rowIndex = 4; rowIndex <= lastRowBeforeExpandingNewWorksheet + 1; rowIndex++)
                {
                    if (rowIndex == lastRowBeforeExpandingNewWorksheet + 1)
                    {
                        newWorksheet.Cells[rowIndex + 1, letterIndex].Value = addedTogether;
                        break;
                    }

                    addedTogether += Convert.ToInt32(newWorksheet.Cells[rowIndex, letterIndex].Value.ToString() ?? "0".Trim('`').Trim());
                }
            }
        }

        // everything added up data
        var lastColumn = newWorksheet.Dimension.End.Column + 1;
        newWorksheet.Cells[1, lastColumn, 3, lastColumn].Merge = true;
        newWorksheet.Cells[1, lastColumn, 3, lastColumn].Value = "Suma";

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
                    newWorksheet.Cells[rowIndex, letterIndex].Value = addedTogether;
                    //Console.WriteLine($"Assigned value '{addedTogether}' to row {rowIndex} and column {letterIndex}");
                    break;
                }

                if ((letterIndex - 1) % totalCategories != 0 && !onlySelectedColumns)
                {
                    addedTogether += Convert.ToInt32(newWorksheet.Cells[rowIndex, letterIndex].Value.ToString() ?? "0".Trim('`').Trim());
                }
                else if ((letterIndex - 1) % totalCategories == 0 && onlySelectedColumns)
                {
                    addedTogether += Convert.ToInt32(newWorksheet.Cells[rowIndex, letterIndex].Value.ToString() ?? "0".Trim('`').Trim());
                }
            }
        }

        TimedLog($"{newWorksheet.Name} | Added additional calculations 80%");

        // styling -------------------------------------------------------------------------------------------------------------------------------------------------

        // 1st to 3rd rows
        newWorksheet.Cells[1, 1, 3, newWorksheet.Dimension.End.Column].Style.Font.Bold = true;
        newWorksheet.Cells[1, 1, 3, newWorksheet.Dimension.End.Column].Style.HorizontalAlignment =
            ExcelHorizontalAlignment.Center;

        newWorksheet.Cells[1, 2, 1, newWorksheet.Dimension.End.Column - 1].Merge = true;

        newWorksheet.Cells["A1:A3"].Merge = true;
        newWorksheet.Cells["A1:A3"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);
        newWorksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        newWorksheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


        for (var columnIndex = 2; columnIndex <= newWorksheet.Dimension.End.Column; columnIndex++)
        {
            if ((columnIndex - 1) % totalCategories == 0)
            {
                newWorksheet.Cells[2, columnIndex - lastVehicleCategory, 2, columnIndex].Style.Border
                    .BorderAround(ExcelBorderStyle.Thick, Color.Black);

                newWorksheet.Cells[3, columnIndex - lastVehicleCategory, 3, columnIndex].Style.Border
                    .BorderAround(ExcelBorderStyle.Thick, Color.Black);
            }
        }

        newWorksheet.Cells[1, newWorksheet.Dimension.End.Column, 3, newWorksheet.Dimension.End.Column].Style.Border
            .BorderAround(ExcelBorderStyle.Thick);

        // time
        newWorksheet.Cells["A1:A" + lastRowBeforeExpandingNewWorksheet].Style.Border
            .BorderAround(ExcelBorderStyle.Thick);

        for (var row = 4; row < lastRowBeforeExpandingNewWorksheet; row++)
        {
            var cellValue = newWorksheet.Cells["A" + row].Value?.ToString() ?? "00:00-00:00";
            var cellValueSplit = cellValue.Split("-");

            var dt = DateTime.Parse(cellValueSplit[1]);
            if (dt.ToString("mm") == "00")
            {
                newWorksheet.Cells[row, 1, row, newWorksheet.Dimension.End.Column - 1].Style.Border.Bottom.Style =
                    ExcelBorderStyle.Thick;
            }
        }

        // added together column bordering
        newWorksheet.Cells[1, 1, lastRowBeforeExpandingNewWorksheet, newWorksheet.Dimension.End.Column - 1].Style
            .Border.BorderAround(ExcelBorderStyle.Thick);

        for (var columnIndex = 2; columnIndex <= newWorksheet.Dimension.End.Column; columnIndex++)
        {
            if ((columnIndex - 1) % totalCategories == 0)
            {
                newWorksheet.Cells[3, columnIndex, lastRowBeforeExpandingNewWorksheet, columnIndex].Style.Border
                    .BorderAround(ExcelBorderStyle.Thick, Color.Black);
            }
        }

        // gray background
        var greyColorRows = lastRowBeforeExpandingNewWorksheet + 2;
        for (var columnIndex = 2; columnIndex <= newWorksheet.Dimension.End.Column; columnIndex++)
        {
            if ((columnIndex - 1) % totalCategories == 0)
            {
                newWorksheet.Cells[3, columnIndex, greyColorRows, columnIndex].Style.Fill.PatternType =
                    ExcelFillStyle.Solid;
                newWorksheet.Cells[3, columnIndex, greyColorRows, columnIndex].Style.Fill.BackgroundColor
                    .SetColor(Color.LightGray);
            }
        }

        // green background
        var greenColorColumns = newWorksheet.Dimension.End.Column - 1;
        for (var rowIndex = 1; rowIndex <= newWorksheet.Dimension.End.Row; rowIndex++)
        {
            newWorksheet.Cells[rowIndex, greenColorColumns].Style.Fill.PatternType = ExcelFillStyle.Solid;
            newWorksheet.Cells[rowIndex, greenColorColumns].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
        }

        // specific column universal settings
        newWorksheet.Cells["A:A"].Style.Font.Bold = true;

        // specific row universal settings
        newWorksheet.Cells["1:3"].Style.Font.Bold = true;

        // universal settings
        newWorksheet.Cells["A1:" + newWorksheet.Dimension.End.Address].Style.Numberformat.Format = "General";
        newWorksheet.Cells["A1:" + newWorksheet.Dimension.End.Address].AutoFitColumns();
        newWorksheet.Cells["A1:" + newWorksheet.Dimension.End.Address].Style.HorizontalAlignment =
            ExcelHorizontalAlignment.Center;
        newWorksheet.Cells["A1:" + newWorksheet.Dimension.End.Address].Style.VerticalAlignment =
            ExcelVerticalAlignment.Center;

        toFormatPackage.Save();
        TimedLog($"{newWorksheet.Name} | Styled and saved worksheet 100%");
    }
}
