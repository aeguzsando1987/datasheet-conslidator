using Serilog;

namespace ConsolidadorReportes.Repositories;

public class DirectoryScanner : IDirectoryScanner
{
    private readonly ILogger _logger;
    private readonly string[] _extensionesValidas = { ".xlsx", ".xlsm", ".xls" };
    private const string PREFIJO_ARCHIVO = "REPORTE SEMANAL";

    public DirectoryScanner(ILogger logger)
    {
        _logger = logger;
    }

    public List<string> EscanearDirectorios(string directorioRaiz)
    {
        _logger.Information("Escaneando directorio raíz: {DirectorioRaiz}", directorioRaiz);

        if (!Directory.Exists(directorioRaiz))
        {
            throw new DirectoryNotFoundException($"El directorio raíz no existe: {directorioRaiz}");
        }

        var archivosValidos = new List<string>();

        // Buscar archivos en todos los subdirectorios
        var todosLosArchivos = Directory.GetFiles(
            directorioRaiz,
            "*.xls*",
            SearchOption.AllDirectories
        );

        foreach (var archivo in todosLosArchivos)
        {
            if (ValidarArchivo(archivo))
            {
                archivosValidos.Add(archivo);
                _logger.Debug("Archivo válido encontrado: {Archivo}", Path.GetFileName(archivo));
            }
        }

        _logger.Information("Total de archivos válidos encontrados: {Cantidad}", archivosValidos.Count);

        return archivosValidos;
    }

    public bool ValidarArchivo(string rutaArchivo)
    {
        var nombreArchivo = Path.GetFileName(rutaArchivo);
        var extension = Path.GetExtension(rutaArchivo).ToLower();

        // Validar que no sea un archivo temporal
        if (nombreArchivo.StartsWith("~$"))
        {
            _logger.Debug("Ignorando archivo temporal: {Archivo}", nombreArchivo);
            return false;
        }

        // Validar que inicie con "REPORTE SEMANAL"
        if (!nombreArchivo.StartsWith(PREFIJO_ARCHIVO, StringComparison.OrdinalIgnoreCase))
        {
            _logger.Debug("Archivo no inicia con '{Prefijo}': {Archivo}", PREFIJO_ARCHIVO, nombreArchivo);
            return false;
        }

        // Validar extensión
        if (!_extensionesValidas.Contains(extension))
        {
            _logger.Debug("Extensión no válida: {Extension} en {Archivo}", extension, nombreArchivo);
            return false;
        }

        // Validar que el archivo sea accesible
        try
        {
            using var _ = File.Open(rutaArchivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return true;
        }
        catch (Exception ex)
        {
            _logger.Warning("No se puede acceder al archivo {Archivo}: {Error}", nombreArchivo, ex.Message);
            return false;
        }
    }
}
