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

public class EmptyArgumentException : Exception
{
    public EmptyArgumentException()
        : base("Some value was detected to be empty when it wasn't supposed to be.")
    {
    }

    public EmptyArgumentException(string exceptionMessage)
        : base(exceptionMessage)
    {
    }

    public EmptyArgumentException(string exceptionMessage, Exception innerException)
        : base(exceptionMessage, innerException)
    {
    }
}

public class CategoryMatchException : Exception
{
    public CategoryMatchException()
        : base("Category has failed to match any existing ones.")
    {
    }

    public CategoryMatchException(string exceptionMessage)
        : base(exceptionMessage)
    {
    }

    public CategoryMatchException(string exceptionMessage, Exception innerException)
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

public class DateTimeConversionException : Exception
{
    public DateTimeConversionException()
        : base("Failed to convert some value to a DateTime object.")
    {
    }

    public DateTimeConversionException(string exceptionMsg)
        : base(exceptionMsg)
    {
    }

    public DateTimeConversionException(string exceptionMsg, Exception exception)
        : base($"Exception Message | {exceptionMsg} \nException | {exception}")
    {
    }
}

public class PrimaryDataValueException : Exception
{
    public PrimaryDataValueException()
        : base("Found incorrect values in primary data.")
    {
    }

    public PrimaryDataValueException(string exceptionMsg)
        : base(exceptionMsg)
    {
    }

    public PrimaryDataValueException(string exceptionMsg, Exception exception)
        : base($"Exception Message | {exceptionMsg} \nException | {exception}")
    {
    }
}
