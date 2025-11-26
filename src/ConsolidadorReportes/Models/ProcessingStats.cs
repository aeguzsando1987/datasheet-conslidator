namespace ConsolidadorReportes.Models;

/// <summary>
/// Estadisticas del proceso de consolidacion
/// </summary>
public class ProcessingStats
{
    public int ArchivosEncontrados { get; set; }
    public int ArchivosExitosos { get; set; }
    public int ArchivosFallidos { get; set; }
    public List<string> ArchivosConError { get; set; } = new();

    public int TotalFilasBaseDatos { get; set; }
    public int NumInicioBaseDatos { get; set; }
    public int NumFinBaseDatos { get; set; }

    public int TotalFilasPlaneacion { get; set; }
    public int NumInicioPlaneacion { get; set; }
    public int NumFinPlaneacion { get; set; }

    public int TotalFilasReporte { get; set; }
    public int NumInicioReporte { get; set; }
    public int NumFinReporte { get; set; }

    public TimeSpan TiempoEjecucion { get; set; }
}
