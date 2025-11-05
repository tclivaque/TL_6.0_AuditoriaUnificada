// Commands/MainCommand.cs
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq; // Necesario para .ToList()
using TL60_RevisionDeTablas.Models;
using TL60_RevisionDeTablas.Services;
using TL60_RevisionDeTablas.UI;

namespace TL60_RevisionDeTablas.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MainCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 1. Inicializar Servicios
                var sheetsService = new GoogleSheetsService();
                var processor = new ScheduleProcessor(doc, sheetsService);

                var elementosData = new List<ElementData>();

                // 2. Obtener todas las tablas
                var allSchedules = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Schedules)
                    .WhereElementIsNotElementType()
                    .Cast<ViewSchedule>();

                // ==========================================================
                // ===== INICIO DE LÓGICA DE BIFURCACIÓN =====
                // ==========================================================

                foreach (ViewSchedule view in allSchedules)
                {
                    if (view == null || view.Definition == null) continue;

                    // PRIMER FILTRO: El nombre de la tabla debe empezar con "C."
                    if (view.Name.StartsWith("C.", StringComparison.OrdinalIgnoreCase))
                    {
                        Parameter groupParam = view.LookupParameter("GRUPO DE VISTA");
                        string groupValue = groupParam?.AsString() ?? string.Empty;

                        // SEGUNDO FILTRO: Revisar "GRUPO DE VISTA"
                        if (groupValue.StartsWith("C.", StringComparison.OrdinalIgnoreCase))
                        {
                            // CASO A: "Tabla de Metrados"
                            // Ejecutar la auditoría completa
                            elementosData.Add(processor.ProcessSingleElement(view));
                        }
                        else
                        {
                            // CASO B: "Tabla de Soporte"
                            // Generar un trabajo de renombrado y corrección de parámetros
                            elementosData.Add(processor.CreateRenamingJob(view));
                        }
                    }
                }
                // ==========================================================
                // ===== FIN DE LÓGICA DE BIFURCACIÓN =====
                // ==========================================================

                // 3. Construir datos de diagnóstico
                var diagnosticBuilder = new DiagnosticDataBuilder();
                List<DiagnosticRow> diagnosticRows = diagnosticBuilder.BuildDiagnosticRows(elementosData);

                // 4. Preparar los Writers Asíncronos
                var writerAsync = new ScheduleUpdateAsync();
                var viewActivator = new ViewActivatorAsync();

                // 5. Crear y mostrar ventana Modeless
                var mainWindow = new MainWindow(
                    diagnosticRows,
                    elementosData,
                    doc,
                    writerAsync,
                    viewActivator
                );

                mainWindow.Show(); // Modeless

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}\nStackTrace: {ex.StackTrace}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }
    }
}