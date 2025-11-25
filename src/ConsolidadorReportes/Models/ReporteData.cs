namespace ConsolidadorReportes.Models;

/// <summary>
/// Contiene los datos de un archivo "REPORTE SEMANAL"
/// </summary>
public class ReporteData
{
    public string ArchivoOrigen { get; set; } = string.Empty;
    public List<BaseDatosRow> BaseDatos { get; set; } = new();
    public List<PlaneacionRow> Planeacion { get; set; } = new();
    public List<ReporteRow> Reporte { get; set; } = new();
}
