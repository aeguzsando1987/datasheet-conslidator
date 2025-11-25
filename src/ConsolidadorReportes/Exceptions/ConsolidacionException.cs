namespace ConsolidadorReportes.Exceptions;

public class ConsolidacionException : Exception
{
    public ConsolidacionException(string message) : base(message)
    {
    }

    public ConsolidacionException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
