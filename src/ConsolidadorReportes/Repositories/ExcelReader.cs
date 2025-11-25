using ClosedXML.Excel;
using Serilog;
using ConsolidadorReportes.Configuration;
using ConsolidadorReportes.Models;

namespace ConsolidadorReportes.Repositories;

public class ExcelReader : IExcelReader
{
    private readonly ILogger _logger;
    private readonly AppSettings _settings;

    public ExcelReader(ILogger logger, AppSettings settings)
    {
        _logger = logger;
        _settings = settings;
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
                RESPONSABLE = row.Cell(2).GetString(),
                REGION = row.Cell(3).GetString(),
                SEMANA = row.Cell(4).GetString(),
                FECHA = row.Cell(5).TryGetValue(out DateTime fecha) ? fecha : null,
                CLASIFICACION = row.Cell(6).GetString(),
                NOMBRE_DE_LA_EMPRESA = row.Cell(7).GetString(),
                GIRO = row.Cell(8).GetString(),
                SECTOR = row.Cell(9).GetString(),
                ESTADO = row.Cell(10).GetString(),
                CIUDAD = row.Cell(11).GetString(),
                DOMICILIO = row.Cell(12).GetString(),
                CONTACTO = row.Cell(13).GetString(),
                PUESTO = row.Cell(14).GetString(),
                EMAIL = row.Cell(15).GetString(),
                TELEFONO = row.Cell(16).GetString(),
                WHATSAPP = row.Cell(17).GetString(),
                FUENTE_DE_INFORMACION = row.Cell(18).GetString(),
                B2B = row.Cell(19).GetString(),
                FECHA_DE_VISITA = row.Cell(20).TryGetValue(out DateTime fechaVisita) ? fechaVisita : null,
                OPORTUNIDAD = row.Cell(21).GetString()
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
                RESPONSABLE = row.Cell(2).GetString(),
                REGION = row.Cell(3).GetString(),
                SEMANA = row.Cell(4).GetString(),
                FECHA = row.Cell(5).TryGetValue(out DateTime fecha) ? fecha : null,
                NOMBRE_DE_LA_EMPRESA = row.Cell(6).GetString(),
                FUENTE_DE_INFORMACION = row.Cell(7).GetString(),
                COMENTARIOS = row.Cell(8).GetString()
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
                RESPONSABLE = row.Cell(2).GetString(),
                REGION = row.Cell(3).GetString(),
                SEMANA = row.Cell(4).GetString(),
                FECHA = row.Cell(5).TryGetValue(out DateTime fecha) ? fecha : null,
                NOMBRE_DE_LA_EMPRESA = row.Cell(6).GetString(),
                FUENTE_DE_INFORMACION = row.Cell(7).GetString(),
                ACTIVIDAD_PROGRAMADA = row.Cell(8).GetString(),
                COMENTARIOS = row.Cell(9).GetString(),
                NECESIDAD_DETECTADA = row.Cell(10).GetString()
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
