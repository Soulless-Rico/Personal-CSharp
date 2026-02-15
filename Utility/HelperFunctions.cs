using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelFormatterConsole.Utility;

public static class HelperFunctions
{
    public static void ErrorLog(string errorMessage)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] <!> {errorMessage} <!>");
        Console.ResetColor();
    }

    public static void WarningLog(string msg)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] <?> {msg} <?>");
        Console.ResetColor();
    }

    public static void BorderAround(ExcelWorksheet ws, string cellCode)
    {
        var cellData = ws.Cells[cellCode];
        BorderAround(ws, cellData.Start.Row, cellData.Start.Column);
    }

    public static void BorderAround(ExcelWorksheet ws, int cellRow, int cellColumn, ExcelBorderStyle style = ExcelBorderStyle.Thin)
    {
        var border = ws.Cells[cellRow, cellColumn].Style.Border;

        border.Bottom.Style = style;
        border.Bottom.Color.SetColor(Color.Black);

        border.Top.Style = style;
        border.Top.Color.SetColor(Color.Black);

        border.Left.Style = style;
        border.Left.Color.SetColor(Color.Black);

        border.Right.Style = style;
        border.Right.Color.SetColor(Color.Black);
    }

    public static void BorderAround(ExcelWorksheet ws, int cellRow, int cellColumn, int cellRowEnd, int cellColumnEnd, ExcelBorderStyle style = ExcelBorderStyle.Thin)
    {
        ApplyBorderIfEmpty(ws.Cells[cellRow, cellColumn, cellRow, cellColumnEnd].Style.Border.Top, style);

        ApplyBorderIfEmpty(ws.Cells[cellRowEnd, cellColumn, cellRowEnd, cellColumnEnd].Style.Border.Bottom, style);

        ApplyBorderIfEmpty(ws.Cells[cellRow, cellColumn, cellRowEnd, cellColumn].Style.Border.Left, style);

        ApplyBorderIfEmpty(ws.Cells[cellRow, cellColumnEnd, cellRowEnd, cellColumnEnd].Style.Border.Right, style);
    }

    public static void BorderEverythingInRange(ExcelWorksheet ws, int rowStart, int columnStart, int rowEnd, int columnEnd, ExcelBorderStyle style = ExcelBorderStyle.Thin)
    {
        var border = ws.Cells[rowStart, columnStart, rowEnd, columnEnd].Style.Border;

        border.Bottom.Style = style;
        border.Bottom.Color.SetColor(Color.Black);

        border.Top.Style = style;
        border.Top.Color.SetColor(Color.Black);

        border.Left.Style = style;
        border.Left.Color.SetColor(Color.Black);

        border.Right.Style = style;
        border.Right.Color.SetColor(Color.Black);
    }

    private static void ApplyBorderIfEmpty(ExcelBorderItem border, ExcelBorderStyle style)
    {
        if (border.Style != ExcelBorderStyle.None) return;
        border.Style = style;
        border.Color.SetColor(Color.Black);
    }

    public static async Task StartSpinner(CancellationToken token)
    {
        char[] spinChars = ['|', '/', '-', '\\'];
        var i = 0;

        while (!token.IsCancellationRequested)
        {
            Console.Write($"\rProcessing Excel Data ... {spinChars[i % 4]} ");
            i++;

            await Task.Delay(100);
        }

        Console.Write("\rDone!                                      \n");
    }

    private static string _logFile = "";

    public static void CreateIfNotFoundLogFolder(string formatName)
    {
        var logDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExcelFormatterLogs"
        );

        Directory.CreateDirectory(logDirectory);

        var logFileTemp = Path.Combine(logDirectory, formatName + ".log");
        File.AppendAllText(logFileTemp, "Log created");

        _logFile = logFileTemp;
    }

    public static void LogToFile(string logMessage)
    {
        if (string.IsNullOrWhiteSpace(_logFile))
        {
            Console.WriteLine("Can't log bacause log file has not been assigned");
            return;
        }

        try
        {
            File.AppendAllText(_logFile, $"[{DateTime.Now:HH:mm:ss}] |: {logMessage} {Environment.NewLine}");
        }
        catch
        {
            Console.WriteLine("Something went wrong when trying to log a process.");
        }
    }
}
