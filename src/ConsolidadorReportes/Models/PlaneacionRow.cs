namespace ConsolidadorReportes.Models;

/// <summary>
/// Representa una fila de la hoja "Planeación"
/// </summary>
public class PlaneacionRow
{
    public int NUM { get; set; }
    public string? RESPONSABLE { get; set; }
    public string? REGION { get; set; }
    public string? SEMANA { get; set; }
    public DateTime? FECHA { get; set; }
    public string? NOMBRE_DE_LA_EMPRESA { get; set; }
    public string? FUENTE_DE_INFORMACION { get; set; }
    public string? COMENTARIOS { get; set; }
}
