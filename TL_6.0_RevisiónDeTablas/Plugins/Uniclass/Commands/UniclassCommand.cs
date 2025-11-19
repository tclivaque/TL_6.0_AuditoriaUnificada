// Plugins/Uniclass/Commands/UniclassCommand.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TL60_RevisionDeTablas.Core;
using TL60_RevisionDeTablas.Models;
using TL60_RevisionDeTablas.Plugins.Uniclass.Services;
using TL60_RevisionDeTablas.Plugins.Uniclass.UI;

namespace TL60_RevisionDeTablas.Plugins.Uniclass.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class UniclassCommand : IExternalCommand
    {
        private const string SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // PASO 1: INICIALIZAR SERVICIOS
                var sheetsService = new GoogleSheetsService();
                sheetsService.Initialize();

                // PASO 2: INICIALIZAR UNICLASS DATA SERVICE
                string docTitleWithoutExt = Path.GetFileNameWithoutExtension(doc.Title);
                var uniclassService = new UniclassDataService(sheetsService, docTitleWithoutExt);
                uniclassService.LoadClassificationData(SPREADSHEET_ID);

                // PASO 3: LEER CATEGORÍAS A AUDITAR
                List<BuiltInCategory> categoriasAuditar = ReadCategoriasFromSheet(sheetsService, SPREADSHEET_ID);

                if (categoriasAuditar.Count == 0)
                {
                    TaskDialog.Show("Advertencia",
                        "No se encontraron categorías para auditar en Google Sheets.\n\n" +
                        "Asegúrese de que la hoja 'ENTRADAS_SCRIPT_2.0 COBIE' existe y tiene categorías en la celda especificada.");
                    return Result.Cancelled;
                }

                // PASO 4: PROCESAR ELEMENTOS
                var auditService = new UniclassAuditService(doc, uniclassService, categoriasAuditar);
                var elementosData = auditService.ProcessElements();

                if (elementosData.Count == 0)
                {
                    TaskDialog.Show("Información", "No se encontraron elementos para auditar.");
                    return Result.Cancelled;
                }

                // PASO 5: MOSTRAR VENTANA
                var uniclassControl = new UniclassPluginControl(uidoc, elementosData);

                var window = new System.Windows.Window
                {
                    Title = "Auditoría de Parámetros Uniclass",
                    Content = uniclassControl,
                    Width = 1600,
                    Height = 700,
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
                };
                window.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error en Auditoría Uniclass",
                    $"Ocurrió un error:\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}");
                message = ex.Message;
                return Result.Failed;
            }
        }

        /// <summary>
        /// Lee las categorías a auditar desde Google Sheets
        /// Hoja: "ENTRADAS_SCRIPT_2.0 COBIE"
        /// Busca la fila que contenga "CATEGORIAS" en columna A
        /// Lee las categorías de columna B (saltos de línea separan categorías)
        /// </summary>
        private List<BuiltInCategory> ReadCategoriasFromSheet(GoogleSheetsService sheetsService, string spreadsheetId)
        {
            var categorias = new List<BuiltInCategory>();

            try
            {
                string sheetName = "ENTRADAS_SCRIPT_2.0 COBIE";
                var data = sheetsService.ReadData(spreadsheetId, $"'{sheetName}'!A:B");

                if (data == null || data.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine($"WARN: No se encontraron datos en {sheetName}");
                    return categorias;
                }

                // Buscar la fila que contenga "CATEGORIAS" en columna A
                int categoriaRowIndex = -1;
                for (int i = 0; i < data.Count; i++)
                {
                    if (data[i].Count > 0)
                    {
                        string cellValue = GoogleSheetsService.GetCellValue(data[i], 0);
                        if (cellValue != null && cellValue.Trim().Equals("CATEGORIAS", StringComparison.OrdinalIgnoreCase))
                        {
                            categoriaRowIndex = i;
                            break;
                        }
                    }
                }

                if (categoriaRowIndex == -1)
                {
                    System.Diagnostics.Debug.WriteLine("WARN: No se encontró la fila 'CATEGORIAS' en columna A");
                    return categorias;
                }

                // Leer el valor de la celda B en esa fila
                if (data[categoriaRowIndex].Count > 1)
                {
                    string categoriasText = GoogleSheetsService.GetCellValue(data[categoriaRowIndex], 1);

                    if (!string.IsNullOrWhiteSpace(categoriasText))
                    {
                        // Separar por saltos de línea
                        var categoriasArray = categoriasText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                        foreach (var catStr in categoriasArray)
                        {
                            string trimmedCat = catStr.Trim();
                            if (string.IsNullOrWhiteSpace(trimmedCat))
                                continue;

                            // Intentar parsear como BuiltInCategory
                            if (Enum.TryParse(trimmedCat, out BuiltInCategory bic))
                            {
                                categorias.Add(bic);
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"WARN: No se pudo parsear la categoría '{trimmedCat}'");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al leer categorías desde Google Sheets: {ex.Message}");
            }

            return categorias;
        }
    }
}
