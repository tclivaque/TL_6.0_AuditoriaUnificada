using System;
using TL60_AuditoriaUnificada.Core;
using System.Collections.Generic;
using System.Linq;
using TL60_AuditoriaUnificada.Plugins.COBie.Models;
using TL60_AuditoriaUnificada.Models;

namespace TL60_AuditoriaUnificada.Plugins.COBie.Services
{
    /// <summary>
    /// Lee la matriz de control "LISTA DE ACT MANT COBIE"
    /// </summary>
    public class ControlMatrixService
    {
        private readonly GoogleSheetsService _sheetsService;
        private const string CONTROL_SHEET_NAME = "LISTA DE ACT MANT COBIE";
        private readonly Dictionary<string, ModelPermissions> _permissions;

        public ControlMatrixService(GoogleSheetsService sheetsService)
        {
            _sheetsService = sheetsService ?? throw new ArgumentNullException(nameof(sheetsService));
            _permissions = new Dictionary<string, ModelPermissions>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Carga la matriz de control desde Google Sheets
        /// </summary>
        public void LoadMatrix(string spreadsheetId)
        {
            try
            {
                var data = _sheetsService.ReadSheet(spreadsheetId, CONTROL_SHEET_NAME);
                if (data == null || data.Count <= 1)
                {
                    throw new Exception($"La hoja de control '{CONTROL_SHEET_NAME}' está vacía o no se encontró.");
                }

                // Procesar filas (saltando encabezado)
                for (int i = 1; i < data.Count; i++)
                {
                    var row = data[i];

                    // ===== Columna E (índice 4) para Título =====
                    string modelTitlesCell = GoogleSheetsService.GetCellValue(row, 4);
                    if (string.IsNullOrWhiteSpace(modelTitlesCell)) continue;

                    // ===== Columnas G a K (índices 6 a 10) para Permisos =====
                    var perms = new ModelPermissions
                    {
                        Facility = GoogleSheetsService.GetCellValue(row, 6), // Col G
                        Floor = GoogleSheetsService.GetCellValue(row, 7),    // Col H
                        Space = GoogleSheetsService.GetCellValue(row, 8),    // Col I
                        Type = GoogleSheetsService.GetCellValue(row, 9),   // Col J
                        Component = GoogleSheetsService.GetCellValue(row, 10)  // Col K
                    };

                    // Manejar celdas con múltiples títulos (separados por salto de línea)
                    var titles = modelTitlesCell.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string title in titles)
                    {
                        // ===== CAMBIO AQUÍ: Quitar extensión .rvt al cargar =====
                        string trimmedTitle = System.IO.Path.GetFileNameWithoutExtension(title.Trim());
                        if (!string.IsNullOrEmpty(trimmedTitle) && !_permissions.ContainsKey(trimmedTitle))
                        {
                            _permissions.Add(trimmedTitle, perms);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al cargar la matriz de control '{CONTROL_SHEET_NAME}': {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Obtiene los permisos para un título de documento específico (ya sin extensión)
        /// </summary>
        public ModelPermissions GetPermissions(string docTitleWithoutExtension)
        {
            if (_permissions.TryGetValue(docTitleWithoutExtension, out var perms))
            {
                return perms;
            }
            // Si no se encuentra, asumir que todo está permitido
            return ModelPermissions.AllowAll();
        }
    }

    public class ModelPermissions
    {
        public string Facility { get; set; }
        public string Floor { get; set; }
        public string Space { get; set; }
        public string Type { get; set; }
        public string Component { get; set; }

        public bool IsAllowed(string groupName)
        {
            string perm;
            switch (groupName.ToUpper())
            {
                case "FACILITY": perm = Facility; break;
                case "FLOOR": perm = Floor; break;
                case "SPACE": perm = Space; break;
                case "TYPE": perm = Type; break;
                case "COMPONENT": perm = Component; break;
                default: perm = "✓"; break;
            }
            return string.IsNullOrWhiteSpace(perm) || perm.Equals("✓");
        }

        public static ModelPermissions AllowAll()
        {
            return new ModelPermissions { Facility = "✓", Floor = "✓", Space = "✓", Type = "✓", Component = "✓" };
        }
    }
}