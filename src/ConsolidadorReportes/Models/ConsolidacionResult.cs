namespace ConsolidadorReportes.Models;

/// <summary>
/// Resultado de la consolidación con información de renumeración
/// </summary>
public class ConsolidacionResult
{
    public List<BaseDatosRow> BaseDatosConsolidada { get; set; } = new();
    public List<PlaneacionRow> PlaneacionConsolidada { get; set; } = new();
    public List<ReporteRow> ReporteConsolidado { get; set; } = new();

    public int NumInicioBaseDatos { get; set; }
    public int NumFinBaseDatos { get; set; }
    public int NumInicioPlaneacion { get; set; }
    public int NumFinPlaneacion { get; set; }
    public int NumInicioReporte { get; set; }
    public int NumFinReporte { get; set; }
}
