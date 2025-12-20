namespace ExcelFormatterConsole.Utility;

public class MissingDataException : Exception
{
    public MissingDataException()
        : base("A critical error has been caught, the program can not continue running and all operations are stopping...")
    {
    }

    public MissingDataException(string exceptionMessage)
        : base(exceptionMessage)
    {
    }

    public MissingDataException(string exceptionMessage, Exception innerException)
        : base(exceptionMessage, innerException)
    {
    }
}

public class MissingDirectionMach : Exception
{
    public MissingDirectionMach()
        : base("No correct match for selected direction has been found")
    {
    }

    public MissingDirectionMach(string exceptionMessage)
        : base(exceptionMessage)
    {
    }

    public MissingDirectionMach(string exceptionMessage, Exception innerException)
        : base(exceptionMessage, innerException)
    {
    }
}

public class ValueConversionException : Exception
{
    public ValueConversionException()
        : base("Value conversion has failed")
    {
    }

    public ValueConversionException(string exceptionMessage)
        : base(exceptionMessage)
    {
    }

    public ValueConversionException(string exceptionMessage, Exception innerException)
        : base(exceptionMessage, innerException)
    {
    }
}

public class MissingWorksheetException : Exception
{
    public MissingWorksheetException()
        : base("Failed to find matching worksheet.")
    {
    }

    public MissingWorksheetException(string exceptionMsg)
        : base(exceptionMsg)
    {
    }

    public MissingWorksheetException(string exceptionMsg, Exception exception)
        : base($"Exception Message | {exceptionMsg} \nException | {exception}")
    {
    }
}