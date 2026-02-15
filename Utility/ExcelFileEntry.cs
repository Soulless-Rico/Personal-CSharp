using System.Diagnostics;
using OfficeOpenXml;

namespace ExcelFormatterConsole.Utility;

public static class ExcelFileEntry
{
    private static readonly Stopwatch Stopwatch = Stopwatch.StartNew();
    private static void TimedLog(string logMessage)
    {
        Stopwatch.Stop();
        Console.WriteLine($"[{Stopwatch.Elapsed.TotalMilliseconds} ms] ----- {logMessage} -----");
        Stopwatch.Restart();
    }

    private static void AskForInput(string message)
    {
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] <<-->> {message} <<-->> ");
    }

    public static (string toFormatExcelFilePath, string generatedExcelFilePath, string fileName) LoadSelectedFiles()
    {
        Stopwatch stopwatch = Stopwatch.StartNew();

        string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelFiles");
        Directory.CreateDirectory(baseDir);

        string newExcelFileDirectory;
        do
        {
            stopwatch.Stop();
            AskForInput("Enter the directory you would like to create the excel file at");
            newExcelFileDirectory = Console.ReadLine()?.Trim().Trim('"')??"";
            stopwatch.Restart();

            if (File.Exists(newExcelFileDirectory))
            {
                HelperFunctions.ErrorLog("Syntax Error: The provided path is a file path not a directory");
                continue;
            }

            if (!Directory.Exists(newExcelFileDirectory))
            {
                HelperFunctions.ErrorLog("Error: No such directory found");
                continue;
            }
            break;

        } while (true);

        TimedLog("Verified directory");

        string fileName;
        using (var package = new ExcelPackage())
        {
            package.Workbook.Worksheets.Add("1");

            string allowedCharacters = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM0123456789-_";
            do
            {
                int flaggedChars = 0;

                stopwatch.Reset();
                AskForInput("Enter the name of the excel file that will get formatted");
                stopwatch.Start();

                fileName = Console.ReadLine()??"";

                foreach (char c in fileName)
                {
                    if (! allowedCharacters.Contains(c) || string.IsNullOrWhiteSpace(fileName))
                    {
                        HelperFunctions.ErrorLog($"Error: File name can only contain these characters '{allowedCharacters}' | illegal character: '{c}'");
                        flaggedChars++;
                        break;
                    }
                }

                if (fileName == "") { HelperFunctions.ErrorLog("Error: File name can not be left empty"); continue; }

                if (flaggedChars > 0 || string.IsNullOrWhiteSpace(fileName)) { continue; }

                if (File.Exists(Path.Combine(newExcelFileDirectory, (fileName + ".xlsx"))))
                {
                    stopwatch.Stop();
                    AskForInput($"File Override: In the current directory theres already a file under the name '{fileName}' \nDo you wish to override this file? (Y/N)");
                    string decision = Console.ReadLine()?.Trim()??"";
                    stopwatch.Start();

                    if (decision.ToLower() != "y") { continue; }
                }
                break;

            } while (true);

            fileName += ".xlsx";
            package.SaveAs(new FileInfo(Path.Combine(newExcelFileDirectory, fileName)));
        }

        TimedLog($"Created excel file '{fileName}'");

        string newExcelFilePath = Path.Combine(newExcelFileDirectory, fileName);
        string unformattedExcelFilePath;
        do
        {
            stopwatch.Reset();
            AskForInput("Enter the file path of the unformatted excel file");
            unformattedExcelFilePath = (Console.ReadLine()?.Trim().Trim('"')) ?? "";
            stopwatch.Start();

            if (string.IsNullOrWhiteSpace(unformattedExcelFilePath))
            {
                HelperFunctions.ErrorLog("Error: File path can not be left empty");
                continue;
            }

            if (!File.Exists(unformattedExcelFilePath))
            {
                HelperFunctions.ErrorLog("Error: Invalid file path");
                continue;
            }

            if (new FileInfo(unformattedExcelFilePath).Extension.ToLower() != ".xlsx")
            {
                HelperFunctions.ErrorLog("Error: Invalid file extension | provided file path doesnt lead to an excel file");
            }

        } while (!File.Exists(unformattedExcelFilePath) || new FileInfo(unformattedExcelFilePath).Extension.ToLower() != ".xlsx");

        TimedLog("Verified unformatted excel file");
        var fileNameWithoutExtension = fileName.Split(".")[0].Trim();
        return (newExcelFilePath, unformattedExcelFilePath, fileNameWithoutExtension);
    }
}
