// Core/UniclassDataService.cs
using System;
using System.Collections.Generic;
using System.Linq;

namespace TL60_RevisionDeTablas.Core
{
    /// <summary>
    /// Servicio compartido para leer la "Matriz UniClass" desde Google Sheets.
    /// Centraliza la lógica de qué Assembly Code es "REVIT" o "MANUAL".
    /// </summary>
    public class UniclassDataService
    {
        private readonly GoogleSheetsService _sheetsService;
        private readonly string _docTitle;
        private readonly Dictionary<string, string> _classificationCache;

        private const string SHEET_ACTIVO = "Matriz UniClass - ACTIVO";
        private const string SHEET_SITIO = "Matriz UniClass - SITIO";

        // (¡MODIFICADO!) Esta lista ahora define los que leen "ACTIVO"
        private readonly List<string> _activoReadCodes = new List<string>
        {
            "469", "470", "471", "472", "473", "474", "475", "476",
            "477", "478", "479", "480", "455", "456"
        };

        private const string _dualReadCode = "200114-CCC02-MO-ES-000410";

        public UniclassDataService(GoogleSheetsService sheetsService, string docTitle)
        {
            _sheetsService = sheetsService ?? throw new ArgumentNullException(nameof(sheetsService));
            _docTitle = docTitle ?? string.Empty;
            _classificationCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Carga los datos desde Google Sheets al caché.
        /// </summary>
        public void LoadClassificationData(string spreadsheetId)
        {
            var sheetsToRead = new List<string>();

            // ==========================================================
            // ===== (¡LÓGICA INVERTIDA SEGÚN TU REQUERIMIENTO!) =====
            // ==========================================================

            // Regla 2: Si es el código dual, lee AMBAS
            if (_docTitle.Contains(_dualReadCode))
            {
                sheetsToRead.Add(SHEET_ACTIVO);
                sheetsToRead.Add(SHEET_SITIO); // SITIO al final para sobrescribir duplicados si es necesario
            }
            // Regla 1: Si contiene "469", "470", etc., lee SOLO ACTIVO
            else if (_activoReadCodes.Any(code => _docTitle.Contains(code)))
            {
                sheetsToRead.Add(SHEET_ACTIVO);
            }
            // Regla 3 (Default): Todos los demás casos, leen SOLO SITIO
            else
            {
                sheetsToRead.Add(SHEET_SITIO);
            }
            // ==========================================================

            foreach (var sheetName in sheetsToRead)
            {
                try
                {
                    // Leer todo el rango para hacerlo dinámico
                    var data = _sheetsService.ReadData(spreadsheetId, $"'{sheetName}'!A:Z");
                    if (data == null || data.Count <= 1)
                    {
                        System.Diagnostics.Debug.WriteLine($"WARN: No se encontraron datos en {sheetName}");
                        continue;
                    }

                    // Leer encabezados (primera fila)
                    var headerRow = data[0];

                    // Buscar columnas por nombre (eliminando espacios)
                    int assemblyCodeIndex = FindColumnIndex(headerRow, "Assembly Code");
                    int metradoTypeIndex = FindColumnIndex(headerRow, "ORIGEN DE METRADO");

                    if (assemblyCodeIndex == -1)
                    {
                        System.Diagnostics.Debug.WriteLine($"ERROR: No se encontró la columna 'Assembly Code' en {sheetName}");
                        continue;
                    }

                    if (metradoTypeIndex == -1)
                    {
                        System.Diagnostics.Debug.WriteLine($"WARN: No se encontró la columna 'ORIGEN DE METRADO' en {sheetName}, se usará 'REVIT' por defecto");
                    }

                    // Procesar filas de datos (saltar encabezados)
                    foreach (var row in data.Skip(1))
                    {
                        if (row.Count <= assemblyCodeIndex) continue;

                        string assemblyCode = GoogleSheetsService.GetCellValue(row, assemblyCodeIndex);

                        if (string.IsNullOrWhiteSpace(assemblyCode))
                            continue;

                        // Leer tipo de metrado (si la columna existe y tiene datos)
                        string metradoType = "REVIT"; // Valor por defecto
                        if (metradoTypeIndex != -1 && row.Count > metradoTypeIndex)
                        {
                            string valorLeido = GoogleSheetsService.GetCellValue(row, metradoTypeIndex);
                            if (!string.IsNullOrWhiteSpace(valorLeido))
                            {
                                metradoType = valorLeido.Trim();
                            }
                        }

                        // Almacenar en caché (si ya existe será sobrescrito)
                        _classificationCache[assemblyCode.Trim()] = metradoType;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error al leer {sheetName}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Busca el índice de una columna por su nombre de encabezado.
        /// Elimina espacios para hacer matching robusto.
        /// </summary>
        /// <param name="headerRow">Fila de encabezados</param>
        /// <param name="columnName">Nombre de la columna a buscar</param>
        /// <returns>Índice de la columna, o -1 si no se encuentra</returns>
        private int FindColumnIndex(IList<object> headerRow, string columnName)
        {
            if (headerRow == null || string.IsNullOrWhiteSpace(columnName))
                return -1;

            // Normalizar el nombre buscado (eliminar espacios y convertir a mayúsculas)
            string normalizedSearchName = columnName.Replace(" ", "").ToUpperInvariant();

            for (int i = 0; i < headerRow.Count; i++)
            {
                string headerValue = GoogleSheetsService.GetCellValue(headerRow, i);

                if (string.IsNullOrWhiteSpace(headerValue))
                    continue;

                // Normalizar el encabezado de la hoja (eliminar espacios y convertir a mayúsculas)
                string normalizedHeaderValue = headerValue.Replace(" ", "").ToUpperInvariant();

                if (normalizedHeaderValue == normalizedSearchName)
                {
                    return i;
                }
            }

            return -1; // No encontrado
        }

        /// <summary>
        /// Obtiene el tipo de metrado (MANUAL o REVIT) para un Assembly Code.
        /// </summary>
        /// <returns>"MANUAL", "REVIT" o "DESCONOCIDO"</returns>
        public string GetScheduleType(string assemblyCode)
        {
            if (string.IsNullOrWhiteSpace(assemblyCode))
                return "DESCONOCIDO";

            if (_classificationCache.TryGetValue(assemblyCode.Trim(), out string type))
            {
                if (type.Equals("MANUAL", StringComparison.OrdinalIgnoreCase))
                    return "MANUAL";

                return "REVIT";
            }

            return "DESCONOCIDO";
        }
    }
}