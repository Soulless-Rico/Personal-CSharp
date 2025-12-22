using OfficeOpenXml;
using ExcelFormatterConsole.Formatting;
using ExcelFormatterConsole.Utility;

namespace ExcelFormatterConsole;

public static class Formatter
{
    public static void Main()
    {
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ----- Program Started -----");

        var (toFormatExcelFilePath, generatedExcelFilePath) = ExcelFileEntry.LoadSelectedFiles();
        using var genPackage = new ExcelPackage(new FileInfo(generatedExcelFilePath));
        using var toFormatPackage = new ExcelPackage(new FileInfo(toFormatExcelFilePath));

        BasicDirectionsClass.LoadPaths(toFormatExcelFilePath, generatedExcelFilePath);

        var genWs =  TvcbClass.FindCorrectWorksheet(genPackage);
        var toFormatWs =  TvcbClass.Prepare(toFormatPackage);

        BasicDirectionsClass.DefaultData();

        int directions = BasicDirectionsClass.FindAllDirections();
        for (int direction = 1; direction <= directions; direction++)
        {
            BasicDirectionsClass.BasicDirections(direction);
        }

        TvcbClass.FormatMeasuredTime(genWs, toFormatWs);
        TvcbClass.FormatVehicleCategories(genWs, toFormatWs);
        TvcbClass.Styling(toFormatWs);

        toFormatPackage.Save();
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ----- Program Finished -----]");
    }

}
