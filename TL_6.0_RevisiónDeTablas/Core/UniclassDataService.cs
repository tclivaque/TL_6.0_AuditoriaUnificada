// Core/UniclassDataService.cs
using System;
using System.Collections.Generic;
using System.Linq;

namespace TL60_AuditoriaUnificada.Core
{
    /// <summary>
    /// Servicio compartido para leer la "Matriz UniClass" desde Google Sheets.
    /// Centraliza la lógica de qué Assembly Code es "REVIT" o "MANUAL".
    /// </summary>
    /// <summary>
    /// Datos de parámetros Uniclass para un Assembly Code
    /// </summary>
    public class UniclassParameterData
    {
        public string AssemblyCode { get; set; }
        public bool ShouldAudit { get; set; } // true si columna UNICLAS = "✓"
        public string EF_Number { get; set; }
        public string EF_Description { get; set; }
        public string Pr_Number { get; set; }
        public string Pr_Description { get; set; }
        public string Ss_Number { get; set; }
        public string Ss_Description { get; set; }
    }

    public class UniclassDataService
    {
        private readonly GoogleSheetsService _sheetsService;
        private readonly string _docTitle;
        private readonly Dictionary<string, string> _classificationCache;
        private readonly Dictionary<string, UniclassParameterData> _uniclassParametersCache;

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
            _uniclassParametersCache = new Dictionary<string, UniclassParameterData>(StringComparer.OrdinalIgnoreCase);
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
                    // Leer todo el rango hasta la columna AZ (52 columnas) para incluir todos los parámetros Uniclass
                    var data = _sheetsService.ReadData(spreadsheetId, $"'{sheetName}'!A:AZ");
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

                    // Buscar columnas de parámetros Uniclass
                    int uniclasIndex = FindColumnIndex(headerRow, "UNICLAS");
                    int efNumberIndex = FindColumnIndex(headerRow, "Classification.Uniclass.EF.Number");
                    int efDescIndex = FindColumnIndex(headerRow, "Classification.Uniclass.EF.Description");
                    int prNumberIndex = FindColumnIndex(headerRow, "Classification.Uniclass.Pr.Number");
                    int prDescIndex = FindColumnIndex(headerRow, "Classification.Uniclass.Pr.Description");
                    int ssNumberIndex = FindColumnIndex(headerRow, "Classification.Uniclass.Ss.Number");
                    int ssDescIndex = FindColumnIndex(headerRow, "Classification.Uniclass.Ss.Description");

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

                        string trimmedAssemblyCode = assemblyCode.Trim();

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

                        // Almacenar en caché de clasificación
                        _classificationCache[trimmedAssemblyCode] = metradoType;

                        // Leer y almacenar datos de parámetros Uniclass
                        var uniclassData = new UniclassParameterData
                        {
                            AssemblyCode = trimmedAssemblyCode,
                            ShouldAudit = false,
                            EF_Number = string.Empty,
                            EF_Description = string.Empty,
                            Pr_Number = string.Empty,
                            Pr_Description = string.Empty,
                            Ss_Number = string.Empty,
                            Ss_Description = string.Empty
                        };

                        // Verificar si debe auditarse (columna UNICLAS)
                        if (uniclasIndex != -1 && row.Count > uniclasIndex)
                        {
                            string uniclasValue = GoogleSheetsService.GetCellValue(row, uniclasIndex);
                            uniclassData.ShouldAudit = uniclasValue?.Trim() == "✓";
                        }

                        // Leer parámetros Uniclass si las columnas existen
                        if (efNumberIndex != -1 && row.Count > efNumberIndex)
                            uniclassData.EF_Number = GoogleSheetsService.GetCellValue(row, efNumberIndex) ?? string.Empty;

                        if (efDescIndex != -1 && row.Count > efDescIndex)
                            uniclassData.EF_Description = GoogleSheetsService.GetCellValue(row, efDescIndex) ?? string.Empty;

                        if (prNumberIndex != -1 && row.Count > prNumberIndex)
                            uniclassData.Pr_Number = GoogleSheetsService.GetCellValue(row, prNumberIndex) ?? string.Empty;

                        if (prDescIndex != -1 && row.Count > prDescIndex)
                            uniclassData.Pr_Description = GoogleSheetsService.GetCellValue(row, prDescIndex) ?? string.Empty;

                        if (ssNumberIndex != -1 && row.Count > ssNumberIndex)
                            uniclassData.Ss_Number = GoogleSheetsService.GetCellValue(row, ssNumberIndex) ?? string.Empty;

                        if (ssDescIndex != -1 && row.Count > ssDescIndex)
                            uniclassData.Ss_Description = GoogleSheetsService.GetCellValue(row, ssDescIndex) ?? string.Empty;

                        // Almacenar en caché de parámetros Uniclass (si ya existe será sobrescrito)
                        _uniclassParametersCache[trimmedAssemblyCode] = uniclassData;
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

        /// <summary>
        /// Obtiene los datos de parámetros Uniclass para un Assembly Code.
        /// </summary>
        /// <returns>UniclassParameterData o null si no se encuentra</returns>
        public UniclassParameterData GetUniclassParameters(string assemblyCode)
        {
            if (string.IsNullOrWhiteSpace(assemblyCode))
                return null;

            if (_uniclassParametersCache.TryGetValue(assemblyCode.Trim(), out UniclassParameterData data))
            {
                return data;
            }

            return null;
        }

        /// <summary>
        /// Verifica si un Assembly Code tiene parámetros Uniclass definidos.
        /// </summary>
        public bool HasUniclassData(string assemblyCode)
        {
            return GetUniclassParameters(assemblyCode) != null;
        }
    }
}