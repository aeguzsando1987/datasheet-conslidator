using ConsolidadorReportes.Configuration;
using ConsolidadorReportes.Repositories;
using ConsolidadorReportes.Services;
using ConsolidadorReportes.UI;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Serilog;

namespace ConsolidadorReportes;

class Program
{
    static async Task<int> Main(string[] args)
    {
        // Configurar Serilog desde appsettings.json
        var configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        Log.Logger = new LoggerConfiguration()
            .ReadFrom.Configuration(configuration)
            .CreateLogger();

        try
        {
            // Mostrar banner de inicio
            ConsoleUI.MostrarBanner();

            Log.Information("Iniciando aplicacion Consolidador de Reportes");

            // Configurar servicios con inyeccion de dependencias
            var services = new ServiceCollection();
            ConfigurarServicios(services, configuration);

            var serviceProvider = services.BuildServiceProvider();

            // Obtener configuracion
            var settings = serviceProvider.GetRequiredService<IOptions<AppSettings>>().Value;

            // Mostrar configuracion
            ConsoleUI.MostrarConfiguracion(
                settings.Paths.DirectorioRaiz,
                settings.Paths.ArchivoMaestro
            );

            // Ejecutar el proceso principal
            var consolidador = serviceProvider.GetRequiredService<IConsolidadorService>();
            var stats = await consolidador.ConsolidarAsync(
                settings.Paths.DirectorioRaiz,
                Path.Combine(settings.Paths.DirectorioRaiz, settings.Paths.ArchivoMaestro)
            );

            // Mostrar resumen final con Spectre.Console
            ConsoleUI.MostrarResumen(stats);

            // Determinar codigo de salida
            if (stats.ArchivosFallidos == 0 || stats.ArchivosExitosos > 0)
            {
                ConsoleUI.MostrarExito("Proceso completado exitosamente");
                Log.Information("Proceso completado exitosamente");
                return 0; // Exito
            }
            else
            {
                ConsoleUI.MostrarError("Proceso completado con errores");
                Log.Error("Proceso completado con errores");
                return 1; // Error
            }
        }
        catch (Exception ex)
        {
            ConsoleUI.MostrarError($"Error fatal: {ex.Message}");
            Log.Fatal(ex, "Error fatal en la aplicacion");
            return 1;
        }
        finally
        {
            Log.CloseAndFlush();
        }
    }

    /// <summary>
    /// Configura todos los servicios con inyeccion de dependencias
    /// </summary>
    static void ConfigurarServicios(IServiceCollection services, IConfiguration configuration)
    {
        // Configuracion
        services.Configure<AppSettings>(configuration.GetSection("AppSettings"));

        // Serilog como singleton
        services.AddSingleton<ILogger>(Log.Logger);

        // Repositories (Transient - nueva instancia por cada solicitud)
        services.AddTransient<IDirectoryScanner, DirectoryScanner>();
        services.AddTransient<IExcelReader, ExcelReader>();
        services.AddTransient<IExcelWriter, ExcelWriter>();
        services.AddTransient<IDataCleaner, DataCleaner>();
        services.AddTransient<IBackupService, BackupService>();

        // Services (Transient)
        services.AddTransient<IConsolidadorService, ConsolidadorService>();
    }
}
