using ClosedXML.Excel;
using ConsolidadorReportes.Configuration;
using ConsolidadorReportes.Exceptions;
using Microsoft.Extensions.Options;
using Serilog;

namespace ConsolidadorReportes.Repositories;

/// <summary>
/// Servicio de limpieza de archivos Excel origen
/// PRESERVA FILA 6 (primera fila con formulas)
/// BORRA SOLO DESDE FILA 7 EN ADELANTE
/// </summary>
public class DataCleaner : IDataCleaner
{
    private readonly AppSettings _settings;
    private readonly ILogger _logger;

    public DataCleaner(
        IOptions<AppSettings> settings,
        ILogger logger)
    {
        _settings = settings.Value;
        _logger = logger;
    }

    /// <summary>
    /// Limpia un archivo Excel preservando la fila 6 con formulas
    /// </summary>
    public async Task LimpiarArchivoAsync(string rutaArchivo)
    {
        _logger.Information("Iniciando limpieza de archivo: {Archivo}", Path.GetFileName(rutaArchivo));

        try
        {
            using var workbook = new XLWorkbook(rutaArchivo);

            // Limpiar cada hoja
            await LimpiarHojaAsync(workbook, _settings.Tablas.NombreHoja1, "BASE DE DATOS");
            await LimpiarHojaAsync(workbook, _settings.Tablas.NombreHoja2, "PLANEACION");
            await LimpiarHojaAsync(workbook, _settings.Tablas.NombreHoja3, "REPORTE");

            // Guardar cambios
            workbook.Save();
            _logger.Information("Archivo limpiado correctamente: {Archivo}", Path.GetFileName(rutaArchivo));
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error al limpiar archivo: {Archivo}", rutaArchivo);
            throw new ConsolidacionException($"Error al limpiar archivo: {rutaArchivo}", ex);
        }
    }

    /// <summary>
    /// Limpia una hoja especifica preservando formulas en fila 6
    /// </summary>
    private async Task LimpiarHojaAsync(XLWorkbook workbook, string nombreHoja, string tipo)
    {
        try
        {
            if (!workbook.Worksheets.TryGetWorksheet(nombreHoja, out var worksheet))
            {
                _logger.Warning("Hoja '{Hoja}' no encontrada, se omite limpieza", nombreHoja);
                return;
            }

            var ultimaFila = worksheet.LastRowUsed()?.RowNumber() ?? 0;

            // Si solo hay encabezados y fila de formulas, no hay nada que borrar
            if (ultimaFila <= _settings.Tablas.FilaPrimeraConFormulas)
            {
                _logger.Information("{Tipo}: No hay datos para borrar (solo encabezados y fila de formulas)", tipo);
                return;
            }

            // CRITICO: Extraer formulas de la fila 6 ANTES de cualquier operacion
            var formulasFila6 = ExtraerFormulas(worksheet, _settings.Tablas.FilaPrimeraConFormulas);

            if (formulasFila6.Count > 0)
            {
                _logger.Debug("Extraidas {Cantidad} formulas de fila {Fila} como respaldo",
                    formulasFila6.Count,
                    _settings.Tablas.FilaPrimeraConFormulas);
            }

            // BORRAR SOLO DESDE FILA 7 EN ADELANTE
            var filasABorrar = ultimaFila - _settings.Tablas.FilaInicioDatosABorrar + 1;

            if (filasABorrar > 0)
            {
                // Borrar desde fila 7 hasta el final
                worksheet.Rows(
                    _settings.Tablas.FilaInicioDatosABorrar,
                    ultimaFila
                ).Delete();

                _logger.Information("{Tipo}: preservada fila {FilaFormulas} ({CantFormulas} formulas), borradas filas {FilaInicio}-{FilaFin}",
                    tipo,
                    _settings.Tablas.FilaPrimeraConFormulas,
                    formulasFila6.Count,
                    _settings.Tablas.FilaInicioDatosABorrar,
                    ultimaFila);
            }

            // VALIDAR QUE LAS FORMULAS SE PRESERVARON
            var formulasIntactas = VerificarFormulasPreservadas(worksheet, _settings.Tablas.FilaPrimeraConFormulas, formulasFila6);

            if (!formulasIntactas)
            {
                // CRITICO! Las formulas se perdieron
                throw new FormulaLostException(
                    $"Las formulas de la fila {_settings.Tablas.FilaPrimeraConFormulas} se perdieron durante la limpieza de hoja '{nombreHoja}'");
            }

            await Task.CompletedTask;
        }
        catch (FormulaLostException)
        {
            // Re-lanzar excepciones criticas de formulas
            throw;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error al limpiar hoja '{Hoja}'", nombreHoja);
            throw;
        }
    }

    /// <summary>
    /// Extrae las formulas de una fila especifica como backup de seguridad
    /// </summary>
    private Dictionary<int, string> ExtraerFormulas(IXLWorksheet worksheet, int numeroFila)
    {
        var formulas = new Dictionary<int, string>();

        try
        {
            var fila = worksheet.Row(numeroFila);
            var ultimaColumna = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

            for (int col = 1; col <= ultimaColumna; col++)
            {
                var celda = fila.Cell(col);

                // Si la celda tiene una formula, guardarla
                if (celda.HasFormula)
                {
                    formulas[col] = celda.FormulaA1;
                    _logger.Debug("Formula en columna {Col}: {Formula}", col, celda.FormulaA1);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Error al extraer formulas de fila {Fila}", numeroFila);
        }

        return formulas;
    }

    /// <summary>
    /// Verifica que las formulas de una fila se hayan preservado correctamente
    /// </summary>
    private bool VerificarFormulasPreservadas(IXLWorksheet worksheet, int numeroFila, Dictionary<int, string> formulasOriginales)
    {
        try
        {
            // Si no habia formulas originales, no hay nada que verificar
            if (formulasOriginales.Count == 0)
            {
                return true;
            }

            var fila = worksheet.Row(numeroFila);
            var formulasIntactas = true;

            foreach (var (columna, formulaOriginal) in formulasOriginales)
            {
                var celda = fila.Cell(columna);

                if (!celda.HasFormula)
                {
                    _logger.Error("Formula perdida en columna {Col} (original: {Formula})",
                        columna,
                        formulaOriginal);
                    formulasIntactas = false;
                }
                else if (celda.FormulaA1 != formulaOriginal)
                {
                    _logger.Warning("Formula modificada en columna {Col}: '{Original}' -> '{Nueva}'",
                        columna,
                        formulaOriginal,
                        celda.FormulaA1);
                }
            }

            if (formulasIntactas)
            {
                _logger.Debug("Todas las formulas de fila {Fila} se preservaron correctamente", numeroFila);
            }

            return formulasIntactas;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error al verificar formulas de fila {Fila}", numeroFila);
            return false;
        }
    }

    /// <summary>
    /// Valida que un archivo solo contenga encabezados y fila de formulas (sin datos)
    /// </summary>
    public async Task<bool> ValidarLimpiezaAsync(string rutaArchivo)
    {
        try
        {
            using var workbook = new XLWorkbook(rutaArchivo);

            var todasLasHojasLimpias = true;

            foreach (var nombreHoja in new[] {
                _settings.Tablas.NombreHoja1,
                _settings.Tablas.NombreHoja2,
                _settings.Tablas.NombreHoja3
            })
            {
                if (workbook.Worksheets.TryGetWorksheet(nombreHoja, out var worksheet))
                {
                    var ultimaFila = worksheet.LastRowUsed()?.RowNumber() ?? 0;

                    // Debe tener solo hasta la fila 6 (encabezados en fila 5, formulas en fila 6)
                    if (ultimaFila > _settings.Tablas.FilaPrimeraConFormulas)
                    {
                        _logger.Warning("Hoja {Hoja} tiene datos mas alla de fila {Fila}",
                            nombreHoja,
                            _settings.Tablas.FilaPrimeraConFormulas);
                        todasLasHojasLimpias = false;
                    }
                }
            }

            await Task.CompletedTask;
            return todasLasHojasLimpias;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error validando limpieza de archivo: {Archivo}", rutaArchivo);
            return false;
        }
    }

    /// <summary>
    /// Restaura formulas en una fila (rollback de emergencia)
    /// </summary>
    private void RestaurarFormulas(IXLWorksheet worksheet, int numeroFila, Dictionary<int, string> formulas)
    {
        try
        {
            var fila = worksheet.Row(numeroFila);

            foreach (var (columna, formula) in formulas)
            {
                var celda = fila.Cell(columna);
                celda.FormulaA1 = formula;
                _logger.Information("Formula restaurada en columna {Col}: {Formula}", columna, formula);
            }

            _logger.Information("Restauradas {Cantidad} formulas en fila {Fila}", formulas.Count, numeroFila);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error al restaurar formulas en fila {Fila}", numeroFila);
            throw;
        }
    }
}
