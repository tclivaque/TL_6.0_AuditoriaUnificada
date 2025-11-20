// Core/GoogleSheetsService.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;

// (ASEGURAR QUE EL NAMESPACE ES .Core)
namespace TL60_AuditoriaUnificada.Core
{
    public class GoogleSheetsService
    {
        private SheetsService _service;

        // CRÍTICO: Usar el nombre de tu archivo JSON
        private static readonly string CREDENTIALS_PATH = Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
            "revitsheetsintegration-89c34b39c2ae.json" // Nombre de tu Contexto
        );

        public GoogleSheetsService()
        {
            InitializeService();
        }

        private void InitializeService()
        {
            try
            {
                GoogleCredential credential;
                using (var stream = new FileStream(CREDENTIALS_PATH, FileMode.Open, FileAccess.Read))
                {
#pragma warning disable CS0618
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(new[] { SheetsService.Scope.Spreadsheets });
#pragma warning restore CS0618
                }

                _service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "TL60RevisionDeTablas"
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al inicializar Google Sheets. Asegúrate que el archivo JSON esté en la carpeta del DLL: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Lee datos de una hoja
        /// </summary>
        public IList<IList<object>> ReadData(string spreadsheetId, string range)
        {
            try
            {
                var request = _service.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = request.Execute();
                return response.Values ?? new List<IList<object>>();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al leer datos de Google Sheets: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Escribe datos en una celda (batch update)
        /// </summary>
        public void WriteData(
            string spreadsheetId,
            Dictionary<string, string> cellValues)
        {
            try
            {
                var dataToWrite = new List<ValueRange>();
                foreach (var kvp in cellValues)
                {
                    var valueRange = new ValueRange
                    {
                        Range = kvp.Key,
                        Values = new List<IList<object>> { new List<object> { kvp.Value } }
                    };
                    dataToWrite.Add(valueRange);
                }

                var batchRequest = new BatchUpdateValuesRequest
                {
                    Data = dataToWrite,
                    ValueInputOption = "RAW"
                };

                var request = _service.Spreadsheets.Values.BatchUpdate(
                    batchRequest,
                    spreadsheetId
                );
                request.Execute();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al escribir datos en Google Sheets: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Obtiene valor de celda (helper)
        /// </summary>
        public static string GetCellValue(IList<object> row, int columnIndex)
        {
            if (row == null || columnIndex >= row.Count)
                return string.Empty;
            return row[columnIndex]?.ToString()?.Trim() ?? string.Empty;
        }

        // ============================================
        // MÉTODOS ADICIONALES PARA PLUGIN COBie
        // ============================================

        /// <summary>
        /// Inicializa el servicio (no-op ya que se inicializa en el constructor)
        /// Usado por: Plugin COBie (compatibilidad)
        /// </summary>
        public void Initialize()
        {
            // El servicio ya está inicializado en el constructor
            // Este método existe solo para compatibilidad con código existente
            if (_service == null)
            {
                InitializeService();
            }
        }

        /// <summary>
        /// Lee una hoja completa por nombre
        /// Usado por: Plugin COBie
        /// </summary>
        public IList<IList<object>> ReadSheet(string spreadsheetId, string sheetName)
        {
            return ReadData(spreadsheetId, sheetName);
        }

        /// <summary>
        /// Parsea un rango de columnas tipo "D:F" y devuelve (columnaInicio, columnaFin)
        /// Usado por: Plugin COBie
        /// </summary>
        public static (int startCol, int endCol) ParseColumnRange(string columnRange)
        {
            if (string.IsNullOrWhiteSpace(columnRange) || !columnRange.Contains(":"))
                throw new ArgumentException("Formato de rango de columnas inválido. Ejemplo esperado: 'D:F'");

            var parts = columnRange.Split(':');
            if (parts.Length != 2)
                throw new ArgumentException("Formato de rango de columnas inválido. Ejemplo esperado: 'D:F'");

            int startCol = ColumnLetterToIndex(parts[0].Trim());
            int endCol = ColumnLetterToIndex(parts[1].Trim());

            return (startCol, endCol);
        }

        /// <summary>
        /// Convierte una letra de columna (ej: "A", "AB") a índice (0-based)
        /// </summary>
        private static int ColumnLetterToIndex(string columnLetter)
        {
            int index = 0;
            for (int i = 0; i < columnLetter.Length; i++)
            {
                index = index * 26 + (char.ToUpper(columnLetter[i]) - 'A' + 1);
            }
            return index - 1; // 0-based
        }
    }
}