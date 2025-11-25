using ClosedXML.Excel;
using Serilog;
using ConsolidadorReportes.Configuration;
using ConsolidadorReportes.Models;

namespace ConsolidadorReportes.Repositories;

public class ExcelWriter : IExcelWriter
{
    private readonly ILogger _logger;
    private readonly AppSettings _settings;

    public ExcelWriter(ILogger logger, AppSettings settings)
    {
        _logger = logger;
        _settings = settings;
    }

    public async Task EscribirMaestroAsync(string rutaMaestro, ConsolidacionResult consolidado)
    {
        await Task.Run(() =>
        {
            XLWorkbook workbook;
            bool archivoNuevo = false;

            if (File.Exists(rutaMaestro))
            {
                workbook = new XLWorkbook(rutaMaestro);
                _logger.Information("Archivo maestro existente abierto (modo APPEND)");
            }
            else
            {
                workbook = new XLWorkbook();
                archivoNuevo = true;
                _logger.Information("Creando archivo maestro nuevo");
            }

            try
            {
                // Escribir BASE DE DATOS
                AgregarDatosHoja(
                    workbook,
                    _settings.Tablas.NombreHoja1,
                    consolidado.BaseDatosConsolidada,
                    archivoNuevo);

                // Escribir PLANEACION
                AgregarDatosHoja(
                    workbook,
                    _settings.Tablas.NombreHoja2,
                    consolidado.PlaneacionConsolidada,
                    archivoNuevo);

                // Escribir REPORTE
                AgregarDatosHoja(
                    workbook,
                    _settings.Tablas.NombreHoja3,
                    consolidado.ReporteConsolidado,
                    archivoNuevo);

                workbook.SaveAs(rutaMaestro);
                _logger.Information("✓ Archivo maestro guardado: {Ruta}", Path.GetFileName(rutaMaestro));
            }
            finally
            {
                workbook.Dispose();
            }
        });
    }

    private void AgregarDatosHoja<T>(XLWorkbook workbook, string nombreHoja, List<T> datos, bool archivoNuevo) where T : class
    {
        IXLWorksheet worksheet;

        if (workbook.Worksheets.TryGetWorksheet(nombreHoja, out var ws))
        {
            worksheet = ws;
        }
        else
        {
            worksheet = workbook.Worksheets.Add(nombreHoja);
            EscribirEncabezados<T>(worksheet);
        }

        if (datos.Count == 0)
        {
            _logger.Warning("No hay datos para escribir en {NombreHoja}", nombreHoja);
            return;
        }

        // Obtener última fila con datos
        var ultimaFila = worksheet.LastRowUsed()?.RowNumber() ?? _settings.Tablas.FilaEncabezados;
        var filaInicio = ultimaFila + 1;

        _logger.Debug("Escribiendo {Cantidad} filas en {NombreHoja} desde fila {Fila}",
            datos.Count, nombreHoja, filaInicio);

        for (int i = 0; i < datos.Count; i++)
        {
            var fila = filaInicio + i;
            EscribirFila(worksheet, fila, datos[i]);
        }

        // Ajustar anchos de columna (solo si es nuevo)
        if (archivoNuevo)
        {
            worksheet.Columns().AdjustToContents();
        }

        _logger.Information("  ✓ {NombreHoja}: {Cantidad} filas escritas", nombreHoja, datos.Count);
    }

    private void EscribirEncabezados<T>(IXLWorksheet worksheet) where T : class
    {
        var properties = typeof(T).GetProperties();
        var filaEncabezados = _settings.Tablas.FilaEncabezados;

        for (int i = 0; i < properties.Length; i++)
        {
            var celda = worksheet.Cell(filaEncabezados, i + 1);
            celda.Value = properties[i].Name;
            celda.Style.Font.Bold = true;
            celda.Style.Fill.BackgroundColor = XLColor.LightGray;
        }
    }

    private void EscribirFila<T>(IXLWorksheet worksheet, int fila, T dato) where T : class
    {
        var properties = typeof(T).GetProperties();

        for (int i = 0; i < properties.Length; i++)
        {
            var valor = properties[i].GetValue(dato);
            var celda = worksheet.Cell(fila, i + 1);

            if (valor != null)
            {
                if (valor is DateTime fechaValor)
                {
                    celda.Value = fechaValor;
                    celda.Style.NumberFormat.Format = "dd/mm/yyyy";
                }
                else
                {
                    celda.Value = valor.ToString();
                }
            }
        }
    }
}
