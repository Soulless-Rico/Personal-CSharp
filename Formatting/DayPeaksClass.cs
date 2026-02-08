using OfficeOpenXml;

namespace ExcelFormatterConsole.Formatting;

public class DayPeaksClass
//  AM, PM Peaks
{
    private static ExcelWorksheet FindWorksheet(ExcelPackage genPackage, string worksheetName)
    {
        return genPackage.Workbook.Worksheets.First(ws => ws.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase));
    }
    public static void GenerateWorksheet (ExcelPackage toFormatPackage, ExcelPackage genPackage)
    {
        List<string> genWsNames = ["am peak class breakdown", "midday peak class breakdown", "pm peak class breakdown"];
        List<string> worksheetNames = ["Doobedňajšia hod.špička", "Obedňajšia hod.špička", "Poobedňajšia hod.špička"];

        var index = 0;
        foreach (var wsName in worksheetNames)
        {
            var toFormatWs = toFormatPackage.Workbook.Worksheets.Add(wsName);
            var genWs = FindWorksheet(genPackage, genWsNames[index++]);

            List<string> mainNavigation = [];
            List<string> secondaryNavigation = [];

            var lastRow = genWs.Dimension.End.Row;
            var lastColumn = genWs.Dimension.End.Column;

            for (var row = 1; row <= lastRow; row++)
            {
                var cellValue = genWs.Cells[row, 1].Value?.ToString() ?? "";
                secondaryNavigation.Add(cellValue);

                if (row > 3) continue;

                for (var column = 2; column <= lastColumn; column++)
                {
                    cellValue = genWs.Cells[row, column].Value?.ToString() ?? "";
                    mainNavigation.Add(cellValue);
                }
            }
        }
    }
}
