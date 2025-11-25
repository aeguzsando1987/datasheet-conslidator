using ConsolidadorReportes.Models;

namespace ConsolidadorReportes.Repositories;

/// <summary>
/// Escribe datos consolidados en el archivo maestro
/// </summary>
public interface IExcelWriter
{
    Task EscribirMaestroAsync(string rutaMaestro, ConsolidacionResult consolidado);
}
