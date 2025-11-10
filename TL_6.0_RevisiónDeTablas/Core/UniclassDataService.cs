// Core/UniclassDataService.cs
using System;
using System.Collections.Generic;
using System.Linq;
using TL60_RevisionDeTablas.Services;

namespace TL60_RevisionDeTablas.Core
{
    /// <summary>
    /// Servicio compartido para leer la "Matriz UniClass" desde Google Sheets.
    /// Centraliza la lógica de qué Assembly Code es "REVIT" o "MANUAL".
    /// Reemplaza a ManualScheduleService.
    /// </summary>
    public class UniclassDataService
    {
        private readonly GoogleSheetsService _sheetsService;
        private readonly string _docTitle;
        private readonly Dictionary<string, string> _classificationCache;

        private const string SHEET_ACTIVO = "Matriz UniClass - ACTIVO";
        private const string SHEET_SITIO = "Matriz UniClass - SITIO";

        // Códigos de proyecto que SÓLO leen de SITIO
        private readonly List<string> _sitioOnlyCodes = new List<string>
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

            if (_sitioOnlyCodes.Any(code => _docTitle.Contains(code)))
            {
                sheetsToRead.Add(SHEET_SITIO);
            }
            else if (_docTitle.Contains(_dualReadCode))
            {
                sheetsToRead.Add(SHEET_SITIO);
                sheetsToRead.Add(SHEET_ACTIVO);
            }
            else
            {
                sheetsToRead.Add(SHEET_ACTIVO); // Default
            }

            foreach (var sheetName in sheetsToRead)
            {
                try
                {
                    // Leer rango D:L (Col D=AC, Col L=Tipo)
                    var data = _sheetsService.ReadData(spreadsheetId, $"'{sheetName}'!D:L");
                    if (data == null || data.Count <= 1) continue;

                    foreach (var row in data.Skip(1))
                    {
                        if (row.Count < 9) continue; // Asegurar que la fila llega hasta la Col L

                        string assemblyCode = GoogleSheetsService.GetCellValue(row, 0); // Col D (Índice 0 en rango D:L)
                        string metradoType = GoogleSheetsService.GetCellValue(row, 8); // Col L (Índice 8 en rango D:L)

                        if (string.IsNullOrWhiteSpace(assemblyCode))
                            continue;

                        // Almacenar el tipo. Si está vacío, se asume "REVIT".
                        _classificationCache[assemblyCode.Trim()] = string.IsNullOrWhiteSpace(metradoType) ? "REVIT" : metradoType.Trim();
                    }
                }
                catch (Exception ex)
                {
                    // No lanzar excepción, solo registrar y continuar.
                    System.Diagnostics.Debug.WriteLine($"Error al leer {sheetName}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Obtiene el tipo de metrado (MANUAL o REVIT) para un Assembly Code.
        /// </summary>
        /// <returns>"MANUAL", "REVIT" o "DESCONOCIDO"</returns>
        public string GetScheduleType(string assemblyCode)
        {
            if (string.IsNullOrWhiteSpace(assemblyCode))
                return "DESCONOCIDO"; // No se puede determinar

            if (_classificationCache.TryGetValue(assemblyCode.Trim(), out string type))
            {
                if (type.Equals("MANUAL", StringComparison.OrdinalIgnoreCase))
                    return "MANUAL";

                return "REVIT"; // Incluye "REVIT" y cualquier otro valor no vacío
            }

            return "DESCONOCIDO"; // No se encontró en la base de datos
        }
    }
}