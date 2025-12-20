using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelFormatterConsole.Utility;

public static class HelperFunctions
{
    public static void ErrorLog(string errorMessage)
    {
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] <!> {errorMessage} <!>");
    }

    public static void BorderAround(ExcelWorksheet ws, string cellCode)
    {
        var border = ws.Cells[cellCode].Style.Border;

        border.Bottom.Style = ExcelBorderStyle.Thin;
        border.Bottom.Color.SetColor(Color.Black);

        border.Top.Style = ExcelBorderStyle.Thin;
        border.Top.Color.SetColor(Color.Black);

        border.Left.Style = ExcelBorderStyle.Thin;
        border.Left.Color.SetColor(Color.Black);

        border.Right.Style = ExcelBorderStyle.Thin;
        border.Right.Color.SetColor(Color.Black);
    }
}