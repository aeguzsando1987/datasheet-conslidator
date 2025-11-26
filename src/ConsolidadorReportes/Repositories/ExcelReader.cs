using ClosedXML.Excel;
using Serilog;
using ConsolidadorReportes.Configuration;
using ConsolidadorReportes.Models;
using Microsoft.Extensions.Options;

namespace ConsolidadorReportes.Repositories;

public class ExcelReader : IExcelReader
{
    private readonly ILogger _logger;
    private readonly AppSettings _settings;

    public ExcelReader(ILogger logger, IOptions<AppSettings> settings)
    {
        _logger = logger;
        _settings = settings.Value;
    }

    /// <summary>
    /// Obtiene el valor de una celda de forma segura sin intentar evaluar formulas
    /// Usa el valor en cache (ultimo valor calculado por Excel)
    /// </summary>
    private string ObtenerValorCelda(IXLCell celda)
    {
        if (celda == null || celda.IsEmpty())
            return string.Empty;

        try
        {
            // Intentar obtener el valor directamente
            if (celda.TryGetValue(out string valor))
            {
                return valor ?? string.Empty;
            }

            // Si tiene formula, usar el valor en cache (ultimo valor calculado)
            if (celda.HasFormula)
            {
                return celda.CachedValue.ToString();
            }

            // Como ultimo recurso, convertir el valor a string
            return celda.Value.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    public ReporteData LeerArchivo(string rutaArchivo)
    {
        _logger.Information("Leyendo archivo: {Archivo}", Path.GetFileName(rutaArchivo));

        var reporteData = new ReporteData { ArchivoOrigen = rutaArchivo };

        try
        {
            using var workbook = new XLWorkbook(rutaArchivo);

            // Leer hoja 1: BASE DE DATOS
            reporteData.BaseDatos = LeerHojaBaseDatos(workbook);
            _logger.Debug("  BASE DE DATOS: {Filas} filas leídas", reporteData.BaseDatos.Count);

            // Leer hoja 2: PLANEACION
            reporteData.Planeacion = LeerHojaPlaneacion(workbook);
            _logger.Debug("  PLANEACION: {Filas} filas leídas", reporteData.Planeacion.Count);

            // Leer hoja 3: REPORTE
            reporteData.Reporte = LeerHojaReporte(workbook);
            _logger.Debug("  REPORTE: {Filas} filas leídas", reporteData.Reporte.Count);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error al leer archivo {Archivo}", rutaArchivo);
            throw;
        }

        return reporteData;
    }

    private List<BaseDatosRow> LeerHojaBaseDatos(XLWorkbook workbook)
    {
        var nombreHoja = _settings.Tablas.NombreHoja1;
        var worksheet = workbook.Worksheets.FirstOrDefault(w =>
            w.Name.Equals(nombreHoja, StringComparison.OrdinalIgnoreCase));

        if (worksheet == null)
        {
            _logger.Warning("Hoja '{NombreHoja}' no encontrada", nombreHoja);
            return new List<BaseDatosRow>();
        }

        var datos = new List<BaseDatosRow>();
        var filaInicio = _settings.Tablas.FilaEncabezados + 1; // Fila 6 (después de encabezados)
        var filaFin = worksheet.LastRowUsed()?.RowNumber() ?? filaInicio;

        for (int fila = filaInicio; fila <= filaFin; fila++)
        {
            var row = worksheet.Row(fila);

            // Verificar si la fila está vacía
            if (row.IsEmpty())
                continue;

            var baseDatos = new BaseDatosRow
            {
                // NUM se ignorará (será renumerado)
                RESPONSABLE = ObtenerValorCelda(row.Cell(2)),
                REGION = ObtenerValorCelda(row.Cell(3)),
                SEMANA = ObtenerValorCelda(row.Cell(4)),
                FECHA = row.Cell(5).TryGetValue(out DateTime fecha) ? fecha : null,
                CLASIFICACION = ObtenerValorCelda(row.Cell(6)),
                NOMBRE_DE_LA_EMPRESA = ObtenerValorCelda(row.Cell(7)),
                GIRO = ObtenerValorCelda(row.Cell(8)),
                SECTOR = ObtenerValorCelda(row.Cell(9)),
                ESTADO = ObtenerValorCelda(row.Cell(10)),
                CIUDAD = ObtenerValorCelda(row.Cell(11)),
                DOMICILIO = ObtenerValorCelda(row.Cell(12)),
                CONTACTO = ObtenerValorCelda(row.Cell(13)),
                PUESTO = ObtenerValorCelda(row.Cell(14)),
                EMAIL = ObtenerValorCelda(row.Cell(15)),
                TELEFONO = ObtenerValorCelda(row.Cell(16)),
                WHATSAPP = ObtenerValorCelda(row.Cell(17)),
                FUENTE_DE_INFORMACION = ObtenerValorCelda(row.Cell(18)),
                B2B = ObtenerValorCelda(row.Cell(19)),
                FECHA_DE_VISITA = row.Cell(20).TryGetValue(out DateTime fechaVisita) ? fechaVisita : null,
                OPORTUNIDAD = ObtenerValorCelda(row.Cell(21))
            };

            datos.Add(baseDatos);
        }

        return datos;
    }

    private List<PlaneacionRow> LeerHojaPlaneacion(XLWorkbook workbook)
    {
        var nombreHoja = _settings.Tablas.NombreHoja2;
        var worksheet = workbook.Worksheets.FirstOrDefault(w =>
            w.Name.Equals(nombreHoja, StringComparison.OrdinalIgnoreCase));

        if (worksheet == null)
        {
            _logger.Warning("Hoja '{NombreHoja}' no encontrada", nombreHoja);
            return new List<PlaneacionRow>();
        }

        var datos = new List<PlaneacionRow>();
        var filaInicio = _settings.Tablas.FilaEncabezados + 1;
        var filaFin = worksheet.LastRowUsed()?.RowNumber() ?? filaInicio;

        for (int fila = filaInicio; fila <= filaFin; fila++)
        {
            var row = worksheet.Row(fila);

            if (row.IsEmpty())
                continue;

            var planeacion = new PlaneacionRow
            {
                RESPONSABLE = ObtenerValorCelda(row.Cell(2)),
                REGION = ObtenerValorCelda(row.Cell(3)),
                SEMANA = ObtenerValorCelda(row.Cell(4)),
                FECHA = row.Cell(5).TryGetValue(out DateTime fecha) ? fecha : null,
                NOMBRE_DE_LA_EMPRESA = ObtenerValorCelda(row.Cell(6)),
                FUENTE_DE_INFORMACION = ObtenerValorCelda(row.Cell(7)),
                COMENTARIOS = ObtenerValorCelda(row.Cell(8))
            };

            datos.Add(planeacion);
        }

        return datos;
    }

    private List<ReporteRow> LeerHojaReporte(XLWorkbook workbook)
    {
        var nombreHoja = _settings.Tablas.NombreHoja3;
        var worksheet = workbook.Worksheets.FirstOrDefault(w =>
            w.Name.Equals(nombreHoja, StringComparison.OrdinalIgnoreCase));

        if (worksheet == null)
        {
            _logger.Warning("Hoja '{NombreHoja}' no encontrada", nombreHoja);
            return new List<ReporteRow>();
        }

        var datos = new List<ReporteRow>();
        var filaInicio = _settings.Tablas.FilaEncabezados + 1;
        var filaFin = worksheet.LastRowUsed()?.RowNumber() ?? filaInicio;

        for (int fila = filaInicio; fila <= filaFin; fila++)
        {
            var row = worksheet.Row(fila);

            if (row.IsEmpty())
                continue;

            var reporte = new ReporteRow
            {
                RESPONSABLE = ObtenerValorCelda(row.Cell(2)),
                REGION = ObtenerValorCelda(row.Cell(3)),
                SEMANA = ObtenerValorCelda(row.Cell(4)),
                FECHA = row.Cell(5).TryGetValue(out DateTime fecha) ? fecha : null,
                NOMBRE_DE_LA_EMPRESA = ObtenerValorCelda(row.Cell(6)),
                FUENTE_DE_INFORMACION = ObtenerValorCelda(row.Cell(7)),
                ACTIVIDAD_PROGRAMADA = ObtenerValorCelda(row.Cell(8)),
                COMENTARIOS = ObtenerValorCelda(row.Cell(9)),
                NECESIDAD_DETECTADA = ObtenerValorCelda(row.Cell(10))
            };

            datos.Add(reporte);
        }

        return datos;
    }

    public int ObtenerUltimoNum(string rutaMaestro, string nombreHoja)
    {
        if (!File.Exists(rutaMaestro))
        {
            _logger.Information("Archivo maestro no existe, último NUM = 0");
            return 0;
        }

        try
        {
            using var workbook = new XLWorkbook(rutaMaestro);
            var worksheet = workbook.Worksheets.FirstOrDefault(w =>
                w.Name.Equals(nombreHoja, StringComparison.OrdinalIgnoreCase));

            if (worksheet == null || worksheet.LastRowUsed() == null)
            {
                _logger.Information("Hoja '{NombreHoja}' vacía, último NUM = 0", nombreHoja);
                return 0;
            }

            var ultimaFila = worksheet.LastRowUsed().RowNumber();
            var filaInicio = _settings.Tablas.FilaEncabezados + 1;

            var numeros = new List<int>();

            for (int fila = filaInicio; fila <= ultimaFila; fila++)
            {
                var celda = worksheet.Cell(fila, 1);
                if (celda.TryGetValue(out int num))
                {
                    numeros.Add(num);
                }
            }

            var ultimoNum = numeros.Any() ? numeros.Max() : 0;
            _logger.Information("Último NUM en {NombreHoja}: {UltimoNum}", nombreHoja, ultimoNum);

            return ultimoNum;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error al leer último NUM de {NombreHoja}", nombreHoja);
            return 0;
        }
    }
}
