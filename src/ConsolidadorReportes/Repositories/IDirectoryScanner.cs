namespace ConsolidadorReportes.Repositories;

/// <summary>
/// Escanea directorios buscando archivos "REPORTE SEMANAL*"
/// </summary>
public interface IDirectoryScanner
{
    /// <summary>
    /// Escanea el directorio raíz y retorna lista de archivos válidos
    /// </summary>
    List<string> EscanearDirectorios(string directorioRaiz);

    /// <summary>
    /// Valida que un archivo cumpla los criterios
    /// </summary>
    bool ValidarArchivo(string rutaArchivo);
}
