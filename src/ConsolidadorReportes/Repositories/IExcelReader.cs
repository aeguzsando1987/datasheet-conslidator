using ConsolidadorReportes.Models;

namespace ConsolidadorReportes.Repositories;

/// <summary>
/// Lee archivos Excel y extrae datos
/// </summary>
public interface IExcelReader
{
    /// <summary>
    /// Lee un archivo Excel completo
    /// </summary>
    ReporteData LeerArchivo(string rutaArchivo);

    /// <summary>
    /// Obtiene el último NUM de una hoja del archivo maestro
    /// </summary>
    int ObtenerUltimoNum(string rutaMaestro, string nombreHoja);
}
