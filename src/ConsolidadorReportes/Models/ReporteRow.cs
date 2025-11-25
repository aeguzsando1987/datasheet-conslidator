namespace ConsolidadorReportes.Models;

/// <summary>
/// Representa una fila de la hoja "Reporte semanal"
/// </summary>
public class ReporteRow
{
    public int NUM { get; set; }
    public string? RESPONSABLE { get; set; }
    public string? REGION { get; set; }
    public string? SEMANA { get; set; }
    public DateTime? FECHA { get; set; }
    public string? NOMBRE_DE_LA_EMPRESA { get; set; }
    public string? FUENTE_DE_INFORMACION { get; set; }
    public string? ACTIVIDAD_PROGRAMADA { get; set; }
    public string? COMENTARIOS { get; set; }
    public string? NECESIDAD_DETECTADA { get; set; }
}