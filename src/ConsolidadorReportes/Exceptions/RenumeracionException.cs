namespace ConsolidadorReportes.Exceptions;

/// <summary>
/// Excepción que se lanza cuando hay un error en la renumeración incremental
/// </summary>
public class RenumeracionException : Exception
{
    public RenumeracionException(string message) : base(message)
    {
    }

    public RenumeracionException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
