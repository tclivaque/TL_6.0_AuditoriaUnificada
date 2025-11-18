using System;
using TL60_RevisionDeTablas.Core;
using System.Collections.Generic;
using System.Linq;
using TL60_RevisionDeTablas.Plugins.COBie.Models;
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Plugins.COBie.Services
{
    /// <summary>
    /// Administrador de configuración desde Google Sheets "ENTRADAS_SCRIPT_2.0 COBIE"
    /// </summary>
    public class ConfigurationManager
    {
        private readonly GoogleSheetsService _sheetsService;
        private const string CONFIG_SHEET_NAME = "ENTRADAS_SCRIPT_2.0 COBIE";

        public ConfigurationManager(GoogleSheetsService sheetsService)
        {
            _sheetsService = sheetsService ?? throw new ArgumentNullException(nameof(sheetsService));
        }

        /// <summary>
        /// Lee la configuración desde Google Sheets
        /// Formato esperado: Columna A = Nombre de variable, Columna B = Valor
        /// </summary>
        public CobieConfig ReadConfiguration(string spreadsheetId)
        {
            try
            {
                var config = new CobieConfig();
                config.SpreadsheetId = spreadsheetId;

                // Leer toda la hoja de configuración
                var data = _sheetsService.ReadSheet(spreadsheetId, CONFIG_SHEET_NAME);

                if (data == null || data.Count == 0)
                {
                    throw new Exception($"La hoja '{CONFIG_SHEET_NAME}' está vacía o no existe.");
                }

                // ===== CAMBIO AQUÍ: Convertir los datos en un diccionario para búsqueda =====
                var configDictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var row in data)
                {
                    if (row.Count < 2) continue; // Necesitamos al menos hasta columna B

                    string variable = GoogleSheetsService.GetCellValue(row, 0); // Columna A
                    string valor = GoogleSheetsService.GetCellValue(row, 1);    // Columna B

                    if (string.IsNullOrWhiteSpace(variable) || configDictionary.ContainsKey(variable)) continue;

                    configDictionary.Add(variable, valor);
                }

                // Asignar valores usando el diccionario (más robusto)
                config.NombreHojaAssemblyCode = configDictionary.GetValueOrDefault("NOMBRE_HOJA_ASSEMBLY_CODE", config.NombreHojaAssemblyCode);
                config.NombreHojaRooms = configDictionary.GetValueOrDefault("NOMBRE_HOJA_ROOMS", config.NombreHojaRooms);
                config.RangoColumnasCOBie = configDictionary.GetValueOrDefault("RANGO_COLUMNAS_COBIE", config.RangoColumnasCOBie);

                if (configDictionary.ContainsKey("SPREADSHEET_ID") && !string.IsNullOrWhiteSpace(configDictionary["SPREADSHEET_ID"]))
                {
                    config.SpreadsheetId = configDictionary["SPREADSHEET_ID"];
                }

                if (configDictionary.ContainsKey("CATEGORIAS"))
                {
                    config.Categorias = ParseCategorias(configDictionary["CATEGORIAS"]);
                }

                // ===== CAMBIO AQUÍ: Leer los nuevos valores por defecto =====
                config.ValorDefectoString = configDictionary.GetValueOrDefault("VALOR_DEFECTO_STRING", config.ValorDefectoString);
                config.ValorDefectoIntegerYesNo = configDictionary.GetValueOrDefault("VALOR_DEFECTO_INTEGER_YESNO", config.ValorDefectoIntegerYesNo);
                config.ValorDefectoInteger = configDictionary.GetValueOrDefault("VALOR_DEFECTO_INTEGER", config.ValorDefectoInteger);
                config.ValorDefectoDouble = configDictionary.GetValueOrDefault("VALOR_DEFECTO_DOUBLE", config.ValorDefectoDouble);

                // Validar configuración
                ValidateConfiguration(config);

                return config;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al leer configuración: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Parsea la lista de categorías desde una celda con saltos de línea
        /// </summary>
        private List<string> ParseCategorias(string categoriasRaw)
        {
            if (string.IsNullOrWhiteSpace(categoriasRaw))
            {
                return new List<string>();
            }

            // Dividir por saltos de línea y limpiar
            var categorias = categoriasRaw
                .Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(c => c.Trim())
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .ToList();

            return categorias;
        }

        /// <summary>
        /// Valida que la configuración tenga todos los valores necesarios
        /// </summary>
        private void ValidateConfiguration(CobieConfig config)
        {
            var errores = new List<string>();

            if (string.IsNullOrWhiteSpace(config.SpreadsheetId))
                errores.Add("SPREADSHEET_ID no está definido");

            if (string.IsNullOrWhiteSpace(config.NombreHojaAssemblyCode))
                errores.Add("NOMBRE_HOJA_ASSEMBLY_CODE no está definido");

            if (string.IsNullOrWhiteSpace(config.NombreHojaRooms))
                errores.Add("NOMBRE_HOJA_ROOMS no está definido");

            if (string.IsNullOrWhiteSpace(config.RangoColumnasCOBie))
                errores.Add("RANGO_COLUMNAS_COBIE no está definido");

            if (config.Categorias == null || config.Categorias.Count == 0)
                errores.Add("CATEGORIAS no está definido o está vacío");

            if (errores.Count > 0)
            {
                throw new Exception(
                    $"Configuración incompleta en '{CONFIG_SHEET_NAME}':\n" +
                    string.Join("\n", errores)
                );
            }
        }

        /// <summary>
        /// Obtiene un resumen de la configuración para mostrar al usuario
        /// </summary>
        public static string GetConfigSummary(CobieConfig config)
        {
            return $"CONFIGURACIÓN CARGADA:\n\n" +
                   $"Hoja Assembly Code: {config.NombreHojaAssemblyCode}\n" +
                   $"Hoja Rooms: {config.NombreHojaRooms}\n" +
                   $"Rango COBie: {config.RangoColumnasCOBie}\n" +
                   $"Categorías: {config.Categorias.Count} categorías definidas\n" +
                   $"Spreadsheet ID: {config.SpreadsheetId.Substring(0, 20)}...";
        }
    }

    // Helper para facilitar la búsqueda en el diccionario
    internal static class DictionaryExtensions
    {
        public static string GetValueOrDefault(this Dictionary<string, string> dict, string key, string defaultValue)
        {
            if (dict.TryGetValue(key, out string value) && !string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
            return defaultValue;
        }
    }
}