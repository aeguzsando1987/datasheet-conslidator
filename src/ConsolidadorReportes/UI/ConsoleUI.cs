using Spectre.Console;
using ConsolidadorReportes.Models;

namespace ConsolidadorReportes.UI;

/// <summary>
/// Clase helper para interfaz de usuario con Spectre.Console
/// </summary>
public static class ConsoleUI
{
    /// <summary>
    /// Muestra el banner de inicio de la aplicacion
    /// </summary>
    public static void MostrarBanner()
    {
        AnsiConsole.Clear();

        var rule = new Rule("[bold cyan]CONSOLIDADOR DE REPORTES SEMANALES v2.0[/]")
        {
            Justification = Justify.Center
        };
        AnsiConsole.Write(rule);

        AnsiConsole.MarkupLine("");
        AnsiConsole.MarkupLine("[dim]Sistema de consolidacion automatizada de reportes Excel[/]");
        AnsiConsole.MarkupLine("[dim]Modo: Renumeracion INCREMENTAL | Preservacion de formulas[/]");
        AnsiConsole.MarkupLine("");
    }

    /// <summary>
    /// Muestra informacion de configuracion inicial
    /// </summary>
    public static void MostrarConfiguracion(string directorioRaiz, string archivoMaestro)
    {
        var table = new Table()
            .Border(TableBorder.Rounded)
            .BorderColor(Color.Grey);

        table.AddColumn("[bold]Configuracion[/]");
        table.AddColumn("[bold]Valor[/]");

        table.AddRow("Directorio raiz", $"[cyan]{directorioRaiz}[/]");
        table.AddRow("Archivo maestro", $"[cyan]{archivoMaestro}[/]");

        AnsiConsole.Write(table);
        AnsiConsole.MarkupLine("");
    }

    /// <summary>
    /// Muestra informacion de ultimo NUM del maestro
    /// </summary>
    public static void MostrarUltimoNum(int baseDatos, int planeacion, int reporte)
    {
        AnsiConsole.MarkupLine("[bold yellow]Ultimo NUM en maestro:[/]");
        AnsiConsole.MarkupLine($"  BASE DE DATOS: [cyan]{baseDatos}[/]");
        AnsiConsole.MarkupLine($"  PLANEACION: [cyan]{planeacion}[/]");
        AnsiConsole.MarkupLine($"  REPORTE: [cyan]{reporte}[/]");
        AnsiConsole.MarkupLine("");
    }

    /// <summary>
    /// Ejecuta una tarea con barra de progreso
    /// </summary>
    public static async Task<T> EjecutarConProgreso<T>(
        string descripcion,
        Func<ProgressTask, Task<T>> accion)
    {
        return await AnsiConsole.Progress()
            .AutoClear(false)
            .Columns(
                new TaskDescriptionColumn(),
                new ProgressBarColumn(),
                new PercentageColumn(),
                new SpinnerColumn())
            .StartAsync(async ctx =>
            {
                var task = ctx.AddTask(descripcion);
                return await accion(task);
            });
    }

    /// <summary>
    /// Muestra tabla de resumen final
    /// </summary>
    public static void MostrarResumen(ProcessingStats stats)
    {
        AnsiConsole.MarkupLine("");
        var rule = new Rule("[bold green]RESUMEN FINAL[/]")
        {
            Justification = Justify.Center
        };
        AnsiConsole.Write(rule);
        AnsiConsole.MarkupLine("");

        // Tabla de archivos procesados
        var tablaArchivos = new Table()
            .Border(TableBorder.Rounded)
            .BorderColor(Color.Green);

        tablaArchivos.AddColumn("[bold]Metrica[/]");
        tablaArchivos.AddColumn(new TableColumn("[bold]Valor[/]").Centered());

        tablaArchivos.AddRow("Archivos encontrados", $"[cyan]{stats.ArchivosEncontrados}[/]");
        tablaArchivos.AddRow("Archivos exitosos", $"[green]{stats.ArchivosExitosos}[/]");

        if (stats.ArchivosFallidos > 0)
        {
            tablaArchivos.AddRow("Archivos fallidos", $"[red]{stats.ArchivosFallidos}[/]");
        }

        AnsiConsole.Write(tablaArchivos);
        AnsiConsole.MarkupLine("");

        // Tabla de datos consolidados
        var tablaDatos = new Table()
            .Border(TableBorder.Rounded)
            .BorderColor(Color.Cyan)
            .Title("[bold cyan]Datos Consolidados[/]");

        tablaDatos.AddColumn("[bold]Hoja[/]");
        tablaDatos.AddColumn(new TableColumn("[bold]Filas[/]").Centered());
        tablaDatos.AddColumn(new TableColumn("[bold]NUM Inicio[/]").Centered());
        tablaDatos.AddColumn(new TableColumn("[bold]NUM Fin[/]").Centered());

        if (stats.TotalFilasBaseDatos > 0)
        {
            tablaDatos.AddRow(
                "BASE DE DATOS",
                $"[cyan]{stats.TotalFilasBaseDatos}[/]",
                $"[yellow]{stats.NumInicioBaseDatos}[/]",
                $"[yellow]{stats.NumFinBaseDatos}[/]"
            );
        }

        if (stats.TotalFilasPlaneacion > 0)
        {
            tablaDatos.AddRow(
                "PLANEACION",
                $"[cyan]{stats.TotalFilasPlaneacion}[/]",
                $"[yellow]{stats.NumInicioPlaneacion}[/]",
                $"[yellow]{stats.NumFinPlaneacion}[/]"
            );
        }

        if (stats.TotalFilasReporte > 0)
        {
            tablaDatos.AddRow(
                "REPORTE",
                $"[cyan]{stats.TotalFilasReporte}[/]",
                $"[yellow]{stats.NumInicioReporte}[/]",
                $"[yellow]{stats.NumFinReporte}[/]"
            );
        }

        AnsiConsole.Write(tablaDatos);
        AnsiConsole.MarkupLine("");

        // Mostrar errores si los hay
        if (stats.ArchivosConError.Any())
        {
            AnsiConsole.MarkupLine("[bold red]Archivos con errores:[/]");
            foreach (var error in stats.ArchivosConError)
            {
                AnsiConsole.MarkupLine($"  [red]- {error}[/]");
            }
            AnsiConsole.MarkupLine("");
        }

        // Tiempo de ejecucion
        var tiempoFormateado = FormatearTiempo(stats.TiempoEjecucion);
        AnsiConsole.MarkupLine($"[bold]Tiempo de ejecucion:[/] [green]{tiempoFormateado}[/]");
        AnsiConsole.MarkupLine("");
    }

    /// <summary>
    /// Muestra mensaje de exito
    /// </summary>
    public static void MostrarExito(string mensaje)
    {
        AnsiConsole.MarkupLine($"[bold green]{mensaje}[/]");
    }

    /// <summary>
    /// Muestra mensaje de error
    /// </summary>
    public static void MostrarError(string mensaje)
    {
        AnsiConsole.MarkupLine($"[bold red]{mensaje}[/]");
    }

    /// <summary>
    /// Muestra mensaje de advertencia
    /// </summary>
    public static void MostrarAdvertencia(string mensaje)
    {
        AnsiConsole.MarkupLine($"[bold yellow]{mensaje}[/]");
    }

    /// <summary>
    /// Formatea un TimeSpan a formato legible
    /// </summary>
    private static string FormatearTiempo(TimeSpan tiempo)
    {
        if (tiempo.TotalMinutes < 1)
        {
            return $"{tiempo.Seconds} segundos";
        }
        else if (tiempo.TotalHours < 1)
        {
            return $"{tiempo.Minutes} minutos {tiempo.Seconds} segundos";
        }
        else
        {
            return $"{tiempo.Hours} horas {tiempo.Minutes} minutos";
        }
    }

    /// <summary>
    /// Muestra un separador visual
    /// </summary>
    public static void MostrarSeparador()
    {
        AnsiConsole.Write(new Rule().RuleStyle("grey dim"));
    }
}
