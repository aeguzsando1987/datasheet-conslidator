namespace ConsolidadorReportes.Exceptions;

/// <summary>
/// Excepción que se lanza cuando se detecta pérdida de fórmulas en fila 6
/// </summary>
public class FormulaLostException : Exception
{
    public FormulaLostException(string message) : base(message)
    {
    }

    public FormulaLostException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
