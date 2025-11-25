namespace ConsolidadorReportes.Repositories;

/// <summary>
/// Limpia datos de archivos origen preservando fórmulas en fila 6
/// </summary>
public interface IDataCleaner
{
    Task LimpiarArchivoAsync(string rutaArchivo);
}
