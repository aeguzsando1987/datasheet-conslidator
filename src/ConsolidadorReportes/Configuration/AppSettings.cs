namespace ConsolidadorReportes.Configuration;

public class AppSettings
{
    public PathsConfig Paths { get; set; } = new();
    public OptionsConfig Options { get; set; } = new();
    public TablasConfig Tablas { get; set; } = new();
    public RenumeracionConfig Renumeracion { get; set; } = new();
    public LoggingConfig Logging { get; set; } = new();
}

public class PathsConfig
{
    public string DirectorioRaiz { get; set; } = string.Empty;
    public string ArchivoMaestro { get; set; } = string.Empty;
    public string DirectorioLogs { get; set; } = string.Empty;
    public string DirectorioBackups { get; set; } = string.Empty;
}

public class OptionsConfig
{
    public bool CrearBackup { get; set; }
    public bool LimpiarArchivosOrigen { get; set; }
    public bool ValidarDuplicados { get; set; }
    public bool IncluirSubdirectorios { get; set; }
}

public class TablasConfig
{
    public string NombreHoja1 { get; set; } = string.Empty;
    public string NombreHoja2 { get; set; } = string.Empty;
    public string NombreHoja3 { get; set; } = string.Empty;
    public int FilaEncabezados { get; set; }
    public int FilaPrimeraConFormulas { get; set; }
    public int FilaInicioDatosABorrar { get; set; }
}

public class RenumeracionConfig
{
    public string Modo { get; set; } = "incremental"; // "incremental" o "desde_cero"
}

public class LoggingConfig
{
    public string Nivel { get; set; } = "Information";
    public int RetenerDias { get; set; }
}
