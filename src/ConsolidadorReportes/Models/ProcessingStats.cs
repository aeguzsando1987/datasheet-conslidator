namespace ConsolidadorReportes.Models;

/// <summary>
/// Estadísticas del proceso de consolidación
/// </summary>
public class ProcessingStats
{
    public int ArchivosEncontrados { get; set; }
    public int ArchivosExitosos { get; set; }
    public int ArchivosFallidos { get; set; }
    public List<string> ArchivosConError { get; set; } = new();

    public int TotalFilasBaseDatos { get; set; }
    public int TotalFilasPlaneacion { get; set; }
    public int TotalFilasReporte { get; set; }

    public TimeSpan TiempoEjecucion { get; set; }
}
