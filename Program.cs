using OfficeOpenXml;
using ExcelFormatterConsole.Formatting;
using ExcelFormatterConsole.Utility;

namespace ExcelFormatterConsole;

public static class Formatter
{
    public static void Main()
    {
        // Preparation
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ----- Program Started -----");

        var (toFormatExcelFilePath, generatedExcelFilePath) = ExcelFileEntry.LoadSelectedFiles();
        using var genPackage = new ExcelPackage(new FileInfo(generatedExcelFilePath));
        using var toFormatPackage = new ExcelPackage(new FileInfo(toFormatExcelFilePath));

        // Basic Directions
        BasicDirectionsClass.DefaultData(genPackage.Workbook.Worksheets[0], toFormatPackage.Workbook.Worksheets[0]);

        var directions = BasicDirectionsClass.FindAllDirections(genPackage);
        for (var direction = 1; direction <= directions; direction++)
        {
            BasicDirectionsClass.BasicDirections(direction, genPackage, toFormatPackage);
        }

        // Total Intensity Rundown
        var genWs =  TirClass.FindCorrectWorksheet(genPackage);
        var toFormatWs =  TirClass.Prepare(toFormatPackage);

        TirClass.FormatMeasuredTime(genWs, toFormatWs);
        var totalVehicleCategoryCount = TirClass.FormatVehicleCategories(genWs, toFormatWs);

        var tirPrimaryDataMapping = TirClass.ReadPrimaryData(genPackage, toFormatPackage);
        TirClass.WritePrimaryData(tirPrimaryDataMapping, toFormatWs);

        var addedUpRowData = TirClass.CalculateAddedUpRowData(toFormatWs);
        var addedUpColumnData = TirClass.CalculateAddedUpColumnData(toFormatWs);

        TirClass.CheckForMatchingResults(toFormatWs, addedUpRowData, addedUpColumnData);
        TirClass.Styling(toFormatWs);

        // Total Volume Class Breakdown
        toFormatWs = TvcbClass.Prepare(toFormatPackage);
        genWs = TvcbClass.FindCorrectWorksheet(genPackage);
        TvcbClass.FormatMeasuredTime(genWs, toFormatWs);
        TvcbClass.Navigation(genWs, toFormatWs, directions);
        TvcbClass.Style(toFormatWs);

        // Program End
        toFormatPackage.Save();
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ----- Program Finished -----]");
    }
}
