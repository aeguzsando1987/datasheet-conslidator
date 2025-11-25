using ConsolidadorReportes.Models;

namespace ConsolidadorReportes.Services;

/// <summary>
/// Servicio principal de consolidación
/// </summary>
public interface IConsolidadorService
{
    Task<ProcessingStats> ConsolidarAsync(string directorioRaiz, string rutaMaestro);
}
