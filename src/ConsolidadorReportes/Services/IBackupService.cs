namespace ConsolidadorReportes.Services;

/// <summary>
/// Servicio para gestión de backups automáticos
/// </summary>
public interface IBackupService
{
    Task<string> CrearBackupMaestroAsync(string rutaMaestro);
    Task<string> CrearBackupOrigenAsync(string rutaOrigen);
    Task LimpiarBackupsAntiguosAsync(int diasRetencion);
}
