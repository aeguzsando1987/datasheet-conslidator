using Serilog;
using ConsolidadorReportes.Configuration;
using ConsolidadorReportes.Models;
using ConsolidadorReportes.Repositories;
using Microsoft.Extensions.Options;

namespace ConsolidadorReportes.Services;

public class ConsolidadorService : IConsolidadorService
{
    private readonly ILogger _logger;
    private readonly AppSettings _settings;
    private readonly IDirectoryScanner _directoryScanner;
    private readonly IExcelReader _excelReader;
    private readonly IBackupService _backupService;
    private readonly IExcelWriter _excelWriter;
    private readonly IDataCleaner _dataCleaner;

    public ConsolidadorService(
        ILogger logger,
        IOptions<AppSettings> settings,
        IDirectoryScanner directoryScanner,
        IExcelReader excelReader,
        IBackupService backupService,
        IExcelWriter excelWriter,
        IDataCleaner dataCleaner)
    {
        _logger = logger;
        _settings = settings.Value;
        _directoryScanner = directoryScanner;
        _excelReader = excelReader;
        _backupService = backupService;
        _excelWriter = excelWriter;
        _dataCleaner = dataCleaner;
    }

    public async Task<ProcessingStats> ConsolidarAsync(string directorioRaiz, string rutaMaestro)
    {
        var inicio = DateTime.Now;
        var stats = new ProcessingStats();

        try
        {
            _logger.Information("========================================");
            _logger.Information("INICIANDO CONSOLIDACIÓN DE REPORTES v2.0");
            _logger.Information("Modo: Renumeración INCREMENTAL");
            _logger.Information("========================================");

            // 1. Escanear archivos
            var archivos = _directoryScanner.EscanearDirectorios(directorioRaiz);
            stats.ArchivosEncontrados = archivos.Count;

            if (archivos.Count == 0)
            {
                _logger.Warning("No se encontraron archivos para procesar");
                return stats;
            }

            // 2. Leer último NUM del maestro
            var ultimoNumBaseDatos = _excelReader.ObtenerUltimoNum(rutaMaestro, _settings.Tablas.NombreHoja1);
            var ultimoNumPlaneacion = _excelReader.ObtenerUltimoNum(rutaMaestro, _settings.Tablas.NombreHoja2);
            var ultimoNumReporte = _excelReader.ObtenerUltimoNum(rutaMaestro, _settings.Tablas.NombreHoja3);

            _logger.Information("----------------------------------------");
            _logger.Information("Último NUM en maestro:");
            _logger.Information("  BASE DE DATOS: {Num}", ultimoNumBaseDatos);
            _logger.Information("  PLANEACION: {Num}", ultimoNumPlaneacion);
            _logger.Information("  REPORTE: {Num}", ultimoNumReporte);
            _logger.Information("----------------------------------------");

            // 3. Crear backup del maestro
            if (_settings.Options.CrearBackup && File.Exists(rutaMaestro))
            {
                await _backupService.CrearBackupMaestroAsync(rutaMaestro);
            }

            // 4. Leer archivos
            var reportes = new List<ReporteData>();
            var archivosExitosos = new List<string>();

            for (int i = 0; i < archivos.Count; i++)
            {
                var archivo = archivos[i];
                _logger.Information("[{Numero}/{Total}] Procesando: {Archivo}",
                    i + 1, archivos.Count, Path.GetFileName(archivo));

                try
                {
                    var reporte = _excelReader.LeerArchivo(archivo);
                    reportes.Add(reporte);
                    archivosExitosos.Add(archivo);
                    stats.ArchivosExitosos++;
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error al procesar archivo");
                    stats.ArchivosFallidos++;
                    stats.ArchivosConError.Add($"{Path.GetFileName(archivo)}: {ex.Message}");
                }
            }

            if (reportes.Count == 0)
            {
                _logger.Error("No se pudo procesar ningún archivo");
                return stats;
            }

            // 5. Consolidar datos con renumeración incremental
            _logger.Information("----------------------------------------");
            _logger.Information("Consolidando datos...");

            var consolidado = ConsolidarDatos(
                reportes,
                ultimoNumBaseDatos,
                ultimoNumPlaneacion,
                ultimoNumReporte);

            stats.TotalFilasBaseDatos = consolidado.BaseDatosConsolidada.Count;
            stats.NumInicioBaseDatos = consolidado.NumInicioBaseDatos;
            stats.NumFinBaseDatos = consolidado.NumFinBaseDatos;

            stats.TotalFilasPlaneacion = consolidado.PlaneacionConsolidada.Count;
            stats.NumInicioPlaneacion = consolidado.NumInicioPlaneacion;
            stats.NumFinPlaneacion = consolidado.NumFinPlaneacion;

            stats.TotalFilasReporte = consolidado.ReporteConsolidado.Count;
            stats.NumInicioReporte = consolidado.NumInicioReporte;
            stats.NumFinReporte = consolidado.NumFinReporte;

            _logger.Information("  BASE DE DATOS: {Filas} filas (NUM {Inicio}-{Fin})",
                stats.TotalFilasBaseDatos, consolidado.NumInicioBaseDatos, consolidado.NumFinBaseDatos);
            _logger.Information("  PLANEACION: {Filas} filas (NUM {Inicio}-{Fin})",
                stats.TotalFilasPlaneacion, consolidado.NumInicioPlaneacion, consolidado.NumFinPlaneacion);
            _logger.Information("  REPORTE: {Filas} filas (NUM {Inicio}-{Fin})",
                stats.TotalFilasReporte, consolidado.NumInicioReporte, consolidado.NumFinReporte);

            // 6. Escribir archivo maestro
            _logger.Information("----------------------------------------");
            _logger.Information("Escribiendo archivo maestro (modo APPEND)...");
            await _excelWriter.EscribirMaestroAsync(rutaMaestro, consolidado);

            // 7. Limpiar archivos origen
            if (_settings.Options.LimpiarArchivosOrigen)
            {
                _logger.Information("----------------------------------------");
                _logger.Information("Limpiando archivos origen...");

                for (int i = 0; i < archivosExitosos.Count; i++)
                {
                    var archivo = archivosExitosos[i];
                    _logger.Information("[{Numero}/{Total}] Limpiando: {Archivo}",
                        i + 1, archivosExitosos.Count, Path.GetFileName(archivo));

                    try
                    {
                        // Backup antes de limpiar
                        if (_settings.Options.CrearBackup)
                        {
                            await _backupService.CrearBackupOrigenAsync(archivo);
                        }

                        await _dataCleaner.LimpiarArchivoAsync(archivo);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Error al limpiar archivo");
                    }
                }
            }

            // 8. Resumen final
            stats.TiempoEjecucion = DateTime.Now - inicio;

            _logger.Information("========================================");
            _logger.Information("RESUMEN FINAL");
            _logger.Information("========================================");
            _logger.Information("Archivos procesados: {Exitosos}/{Total}", stats.ArchivosExitosos, stats.ArchivosEncontrados);
            _logger.Information("Archivos con errores: {Fallidos}", stats.ArchivosFallidos);

            if (stats.ArchivosConError.Any())
            {
                _logger.Warning("Errores:");
                foreach (var error in stats.ArchivosConError)
                {
                    _logger.Warning("  - {Error}", error);
                }
            }

            _logger.Information("");
            _logger.Information("Tiempo de ejecución: {Tiempo}", stats.TiempoEjecucion);
            _logger.Information("========================================");
        }
        catch (Exception ex)
        {
            _logger.Fatal(ex, "Error crítico en consolidación");
            throw;
        }

        return stats;
    }

    private ConsolidacionResult ConsolidarDatos(
        List<ReporteData> reportes,
        int ultimoNumBaseDatos,
        int ultimoNumPlaneacion,
        int ultimoNumReporte)
    {
        var resultado = new ConsolidacionResult();

        // Consolidar BASE DE DATOS
        var todasBaseDatos = reportes.SelectMany(r => r.BaseDatos).ToList();
        resultado.BaseDatosConsolidada = RenumerarIncremental(todasBaseDatos, ultimoNumBaseDatos);
        resultado.NumInicioBaseDatos = ultimoNumBaseDatos + 1;
        resultado.NumFinBaseDatos = ultimoNumBaseDatos + todasBaseDatos.Count;

        // Consolidar PLANEACION
        var todasPlaneacion = reportes.SelectMany(r => r.Planeacion).ToList();
        resultado.PlaneacionConsolidada = RenumerarIncremental(todasPlaneacion, ultimoNumPlaneacion);
        resultado.NumInicioPlaneacion = ultimoNumPlaneacion + 1;
        resultado.NumFinPlaneacion = ultimoNumPlaneacion + todasPlaneacion.Count;

        // Consolidar REPORTE
        var todosReporte = reportes.SelectMany(r => r.Reporte).ToList();
        resultado.ReporteConsolidado = RenumerarIncremental(todosReporte, ultimoNumReporte);
        resultado.NumInicioReporte = ultimoNumReporte + 1;
        resultado.NumFinReporte = ultimoNumReporte + todosReporte.Count;

        return resultado;
    }

    /// <summary>
    /// Renumera incrementalmente desde ultimo_num + 1
    /// </summary>
    private List<T> RenumerarIncremental<T>(List<T> datos, int ultimoNum) where T : class
    {
        int nuevoNum = ultimoNum + 1;

        foreach (var item in datos)
        {
            var prop = item.GetType().GetProperty("NUM");
            prop?.SetValue(item, nuevoNum++);
        }

        _logger.Debug("Renumeración: {Cantidad} registros desde NUM {Inicio} hasta {Fin}",
            datos.Count, ultimoNum + 1, nuevoNum - 1);

        return datos;
    }
}
