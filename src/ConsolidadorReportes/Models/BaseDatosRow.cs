namespace ConsolidadorReportes.Models;

/// <summary>
/// Representa una fila de la hoja "BASE DE DATOS DE PROSPECTOS LOCALIZADOS"
/// </summary>
public class BaseDatosRow
{
    public int NUM { get; set; }
    public string? RESPONSABLE { get; set; }
    public string? REGION { get; set; }
    public string? SEMANA { get; set; }
    public DateTime? FECHA { get; set; }
    public string? CLASIFICACION { get; set; }
    public string? NOMBRE_DE_LA_EMPRESA { get; set; }
    public string? GIRO { get; set; }
    public string? SECTOR { get; set; }
    public string? ESTADO { get; set; }
    public string? CIUDAD { get; set; }
    public string? DOMICILIO { get; set; }
    public string? CONTACTO { get; set; }
    public string? PUESTO { get; set; }
    public string? EMAIL { get; set; }
    public string? TELEFONO { get; set; }
    public string? WHATSAPP { get; set; }
    public string? FUENTE_DE_INFORMACION { get; set; }
    public string? B2B { get; set; }
    public DateTime? FECHA_DE_VISITA { get; set; }
    public string? OPORTUNIDAD { get; set; }
}
