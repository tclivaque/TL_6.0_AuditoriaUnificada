using System;
using TL60_RevisionDeTablas.Core;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using TL60_RevisionDeTablas.Plugins.COBie.Models;
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Plugins.COBie.Services
{
    /// <summary>
    /// Lee las definiciones de parámetros desde Google Sheets "PARÁMETROS COBIE"
    /// </summary>
    public class ParameterDefinitionReader
    {
        private readonly GoogleSheetsService _sheetsService;
        private const string PARAMETERS_SHEET_NAME = "PARÁMETROS COBIE";

        // Índices de columnas (0-based)
        private const int COL_GRUPO = 0;              // Columna A
        private const int COL_NOMBRE_PARAMETRO = 1;   // Columna B
        private const int COL_FECHA_TECNICA = 2;      // Columna C
        private const int COL_VALOR_FIJO = 3;         // Columna D
        private const int COL_DESCRIPCION = 4;        // Columna E
        private const int COL_ESTADO = 5;             // Columna F

        public ParameterDefinitionReader(GoogleSheetsService sheetsService)
        {
            _sheetsService = sheetsService ?? throw new ArgumentNullException(nameof(sheetsService));
        }

        /// <summary>
        /// Lee todas las definiciones de parámetros desde Google Sheets
        /// </summary>
        public List<ParameterDefinition> ReadParameterDefinitions(string spreadsheetId)
        {
            try
            {
                var definitions = new List<ParameterDefinition>();

                // Leer toda la hoja
                var data = _sheetsService.ReadSheet(spreadsheetId, PARAMETERS_SHEET_NAME);

                if (data == null || data.Count == 0)
                {
                    throw new Exception($"La hoja '{PARAMETERS_SHEET_NAME}' está vacía o no existe.");
                }

                // Saltar la primera fila (headers)
                for (int i = 1; i < data.Count; i++)
                {
                    var row = data[i];

                    // Verificar que la fila tenga datos mínimos
                    if (row.Count < 2) continue;

                    string grupo = GoogleSheetsService.GetCellValue(row, COL_GRUPO);
                    string nombreParametro = GoogleSheetsService.GetCellValue(row, COL_NOMBRE_PARAMETRO);

                    // Saltar filas vacías o sin nombre de parámetro
                    if (string.IsNullOrWhiteSpace(nombreParametro)) continue;

                    // 🔥 LEER VALOR FIJO Y CONVERTIR FECHAS
                    string valorFijo = GoogleSheetsService.GetCellValue(row, COL_VALOR_FIJO);
                    valorFijo = ConvertirFechaAFormatoISO(valorFijo);

                    var definition = new ParameterDefinition
                    {
                        Grupo = grupo,
                        NombreParametro = nombreParametro,
                        FechaTecnica = GoogleSheetsService.GetCellValue(row, COL_FECHA_TECNICA),
                        ValorFijo = valorFijo,
                        Descripcion = GoogleSheetsService.GetCellValue(row, COL_DESCRIPCION),
                        Estado = GoogleSheetsService.GetCellValue(row, COL_ESTADO)
                    };

                    definitions.Add(definition);
                }

                return definitions;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al leer definiciones de parámetros: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 🔥 SOLUCIÓN: Convierte fechas del formato de Google Sheets al formato ISO COBie
        /// </summary>
        /// <param name="valor">Valor que puede ser una fecha en formato dd/MM/yyyy HH:mm:ss</param>
        /// <returns>Fecha en formato yyyy-MM-ddTHH:mm:ss o el valor original si no es fecha</returns>
        private string ConvertirFechaAFormatoISO(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
                return valor;

            // Formatos de fecha que Google Sheets puede devolver
            string[] formatosGoogleSheets = new[]
            {
                "dd/MM/yyyy HH:mm:ss",  // 30/11/2025 00:00:01
                "dd/MM/yyyy",           // 30/11/2025
                "d/M/yyyy HH:mm:ss",    // 30/11/2025 00:00:01 (sin ceros)
                "d/M/yyyy",             // 30/11/2025 (sin ceros)
            };

            DateTime fecha;

            // Intentar parsear con los formatos de Google Sheets
            if (DateTime.TryParseExact(
                valor,
                formatosGoogleSheets,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out fecha))
            {
                // ✅ CONVERTIR A FORMATO ISO COBie: yyyy-MM-ddTHH:mm:ss
                return fecha.ToString("yyyy-MM-dd'T'HH:mm:ss");
            }

            // Si no es una fecha reconocida, devolver el valor original
            return valor;
        }

        /// <summary>
        /// Agrupa las definiciones por grupo COBie
        /// </summary>
        public Dictionary<string, List<ParameterDefinition>> GroupByGrupo(List<ParameterDefinition> definitions)
        {
            return definitions
                .GroupBy(d => d.Grupo.ToUpper().Trim())
                .ToDictionary(
                    g => g.Key,
                    g => g.ToList()
                );
        }

        /// <summary>
        /// Obtiene solo los parámetros de un grupo específico
        /// </summary>
        public List<ParameterDefinition> GetParametersByGroup(
            List<ParameterDefinition> definitions,
            string grupo)
        {
            return definitions
                .Where(d => d.Grupo.Equals(grupo, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        /// <summary>
        /// Obtiene solo los parámetros con valores fijos
        /// </summary>
        public List<ParameterDefinition> GetParametersWithFixedValues(List<ParameterDefinition> definitions)
        {
            return definitions
                .Where(d => d.TieneValorFijo)
                .ToList();
        }

        /// <summary>
        /// Obtiene solo los parámetros dinámicos (sin valor fijo)
        /// </summary>
        public List<ParameterDefinition> GetDynamicParameters(List<ParameterDefinition> definitions)
        {
            return definitions
                .Where(d => !d.TieneValorFijo)
                .ToList();
        }

        /// <summary>
        /// Busca una definición específica por nombre de parámetro
        /// </summary>
        public ParameterDefinition FindParameter(
            List<ParameterDefinition> definitions,
            string nombreParametro)
        {
            return definitions
                .FirstOrDefault(d => d.NombreParametro.Equals(
                    nombreParametro,
                    StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Genera un resumen de las definiciones cargadas
        /// </summary>
        public static string GetDefinitionsSummary(List<ParameterDefinition> definitions)
        {
            var grouped = definitions.GroupBy(d => d.Grupo).ToList();

            var summary = $"DEFINICIONES DE PARÁMETROS CARGADAS:\n\n";
            summary += $"Total de parámetros: {definitions.Count}\n\n";

            foreach (var group in grouped.OrderBy(g => g.Key))
            {
                int conValorFijo = group.Count(d => d.TieneValorFijo);
                int dinamicos = group.Count() - conValorFijo;

                summary += $"• {group.Key}: {group.Count()} parámetros ";
                summary += $"({conValorFijo} fijos, {dinamicos} dinámicos)\n";
            }

            return summary;
        }
    }
}