using Serilog;
using ConsolidadorReportes.Configuration;
using Microsoft.Extensions.Options;

namespace ConsolidadorReportes.Services;

public class BackupService : IBackupService
{
    private readonly ILogger _logger;
    private readonly AppSettings _settings;

    public BackupService(ILogger logger, IOptions<AppSettings> settings)
    {
        _logger = logger;
        _settings = settings.Value;
    }

    public async Task<string> CrearBackupMaestroAsync(string rutaMaestro)
    {
        return await Task.Run(() =>
        {
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var nombreBackup = $"maestro_backup_{timestamp}.xlsx";
            var rutaBackup = Path.Combine(_settings.Paths.DirectorioBackups, nombreBackup);

            Directory.CreateDirectory(_settings.Paths.DirectorioBackups);

            File.Copy(rutaMaestro, rutaBackup, overwrite: true);

            _logger.Information("Backup maestro creado: {Backup}", nombreBackup);

            return rutaBackup;
        });
    }

    public async Task<string> CrearBackupOrigenAsync(string rutaOrigen)
    {
        return await Task.Run(() =>
        {
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var carpetaBackup = Path.Combine(_settings.Paths.DirectorioBackups, $"origenes_{timestamp}");

            Directory.CreateDirectory(carpetaBackup);

            var nombreArchivo = Path.GetFileName(rutaOrigen);
            var rutaBackup = Path.Combine(carpetaBackup, nombreArchivo);

            File.Copy(rutaOrigen, rutaBackup, overwrite: true);

            _logger.Debug("Backup origen creado: {Backup}", nombreArchivo);

            return rutaBackup;
        });
    }

    public async Task LimpiarBackupsAntiguosAsync(int diasRetencion)
    {
        await Task.Run(() =>
        {
            var directorioBackups = _settings.Paths.DirectorioBackups;

            if (!Directory.Exists(directorioBackups))
                return;

            var fechaLimite = DateTime.Now.AddDays(-diasRetencion);
            var archivosEliminados = 0;

            foreach (var archivo in Directory.GetFiles(directorioBackups))
            {
                var fechaCreacion = File.GetCreationTime(archivo);

                if (fechaCreacion < fechaLimite)
                {
                    File.Delete(archivo);
                    archivosEliminados++;
                }
            }

            if (archivosEliminados > 0)
            {
                _logger.Information("Backups antiguos eliminados: {Cantidad}", archivosEliminados);
            }
        });
    }
}
