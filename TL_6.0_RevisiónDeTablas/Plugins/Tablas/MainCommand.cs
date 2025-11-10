// Plugins/Tablas/MainCommand.cs
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using TL60_RevisionDeTablas.Core;       // <-- Actualizado
using TL60_RevisionDeTablas.Models;
using TL60_RevisionDeTablas.Plugins.Tablas; // <-- Actualizado
using TL60_RevisionDeTablas.Services;
using TL60_RevisionDeTablas.UI;

namespace TL60_RevisionDeTablas.Plugins.Tablas // <-- Actualizado
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MainCommand : IExternalCommand
    {
        private const string SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";

        private static readonly List<string> NOMBRES_WIP = new List<string>
        {
            "TL", "TITO", "PDONTADENEA", "ANDREA", "EFRAIN",
            "PROYECTOSBIM", "ASISTENTEBIM", "LUIS", "DIEGO", "JORGE", "MIGUEL"
        };

        private static readonly Regex _acRegex = new Regex(@"^C\.(\d{2,3}\.)+\d{2,3}");

        // TODO: Definir la lista de categorías para la auditoría de elementos
        private static readonly List<BuiltInCategory> _categoriesToAudit = new List<BuiltInCategory>
        {
            // Por ejemplo:
            // BuiltInCategory.OST_Walls,
            // BuiltInCategory.OST_Floors,
            // BuiltInCategory.OST_Ceilings,
            // BuiltInCategory.OST_Roofs,
            // BuiltInCategory.OST_Columns,
            // BuiltInCategory.OST_StructuralFraming,
            // BuiltInCategory.OST_StructuralFoundation,
            // BuiltInCategory.OST_MechanicalEquipment,
            // BuiltInCategory.OST_DuctTerminal,
            // BuiltInCategory.OST_PipeFixtures,
            // BuiltInCategory.OST_PlumbingFixtures,
            // BuiltInCategory.OST_LightingFixtures,
            // BuiltInCategory.OST_ElectricalEquipment,
            // BuiltInCategory.OST_CableTray,
            // BuiltInCategory.OST_Conduit
        };


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
                string docTitle = Path.GetFileNameWithoutExtension(doc.Title);

                // (MODIFICADO) Usar el nuevo servicio compartido
                var uniclassService = new UniclassDataService(sheetsService, docTitle);
                uniclassService.LoadClassificationData(SPREADSHEET_ID);

                var processor = new ScheduleProcessor(doc, sheetsService, uniclassService, SPREADSHEET_ID);

                var elementosData = new List<ElementData>();
                var existingMetradoCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // 2. Obtener todas las tablas
                var allSchedules = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Schedules)
                    .WhereElementIsNotElementType()
                    .Cast<ViewSchedule>();

                // 3. Procesar Tablas de Planificación (Lógica existente)
                foreach (ViewSchedule view in allSchedules)
                {
                    // ... (Lógica de bifurcación sin cambios)...
                    if (view == null || view.Definition == null) continue;
                    string viewName = view.Name;
                    string viewNameUpper = viewName.ToUpper();

                    if (viewNameUpper.StartsWith("C."))
                    {
                        if (viewNameUpper.Contains("COPY") || viewNameUpper.Contains("COPIA"))
                        {
                            // ... (lógica de COPIA sin cambios) ...
                            string grupoVista = view.LookupParameter("GRUPO DE VISTA")?.AsString() ?? string.Empty;
                            string subGrupoVista = view.LookupParameter("SUBGRUPO DE VISTA")?.AsString() ?? string.Empty;
                            if (grupoVista.Equals("REVISAR", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(subGrupoVista))
                            {
                                continue;
                            }
                            elementosData.Add(processor.CreateCopyReclassifyJob(view));
                            continue;
                        }

                        if (NOMBRES_WIP.Any(name => viewNameUpper.Contains(name)))
                        {
                            elementosData.Add(processor.CreateWipReclassifyJob(view));
                            continue;
                        }

                        Match acMatch = _acRegex.Match(viewName);
                        string assemblyCode = acMatch.Success ? acMatch.Value : "INVALID_AC";

                        // (MODIFICADO) Usar el nuevo servicio
                        string scheduleType = uniclassService.GetScheduleType(assemblyCode);

                        if (scheduleType.Equals("MANUAL", StringComparison.OrdinalIgnoreCase))
                        {
                            // ... (lógica MANUAL sin cambios) ...
                            string grupoVista = view.LookupParameter("GRUPO DE VISTA")?.AsString() ?? string.Empty;
                            string subGrupoVista = view.LookupParameter("SUBGRUPO DE VISTA")?.AsString() ?? string.Empty;
                            if (grupoVista.Equals("REVISAR", StringComparison.OrdinalIgnoreCase) && subGrupoVista.Equals("METRADO MANUAL", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }
                            elementosData.Add(processor.CreateManualRenamingJob(view));
                            continue;
                        }

                        // Es "REVIT" o "DESCONOCIDO"
                        Parameter groupParam = view.LookupParameter("GRUPO DE VISTA");
                        string groupValue = groupParam?.AsString() ?? string.Empty;

                        if (groupValue.StartsWith("C.", StringComparison.OrdinalIgnoreCase))
                        {
                            // CASO A: "Tabla de Metrados" -> Auditar
                            elementosData.Add(processor.ProcessSingleElement(view, assemblyCode));

                            // (NUEVO) Guardar el AC para la auditoría de elementos
                            if (assemblyCode != "INVALID_AC")
                            {
                                existingMetradoCodes.Add(assemblyCode);
                            }
                        }
                        else
                        {
                            // CASO B: "Tabla de Soporte" -> Renombrar
                            elementosData.Add(processor.CreateRenamingJob(view));
                        }
                    }
                }

                // ==========================================================
                // ===== 4. (¡NUEVO!) EJECUTAR AUDITORÍA DE ELEMENTOS (Tablas Faltantes)
                // ==========================================================
                if (_categoriesToAudit.Count > 0)
                {
                    var elementAuditor = new MissingScheduleAuditor(doc, uniclassService);
                    ElementData missingSchedulesReport = elementAuditor.FindMissingSchedules(existingMetradoCodes, _categoriesToAudit);

                    if (missingSchedulesReport != null)
                    {
                        elementosData.Add(missingSchedulesReport);
                    }
                }
                else
                {
                    // Reportar que la auditoría de elementos no está configurada
                    elementosData.Add(new ElementData
                    {
                        ElementId = ElementId.InvalidElementId,
                        Nombre = "Auditoría de Elementos",
                        Categoria = "Sistema",
                        DatosCompletos = false,
                        AuditResults = new List<AuditItem>
                        {
                            new AuditItem
                            {
                                AuditType = "CONFIGURACIÓN",
                                Estado = EstadoParametro.Vacio, // Advertencia
                                Mensaje = "La auditoría de elementos modelados (Tablas Faltantes) no está configurada. " +
                                          "Se debe definir la lista de categorías a revisar en MainCommand.cs."
                            }
                        }
                    });
                }

                // ==========================================================
                // ===== 5. POST-PROCESO: Auditoría de Duplicados
                // ==========================================================
                RunDuplicateCheck(elementosData);

                // 6. Construir datos de diagnóstico
                var diagnosticBuilder = new DiagnosticDataBuilder();
                List<DiagnosticRow> diagnosticRows = diagnosticBuilder.BuildDiagnosticRows(elementosData);

                // 7. Preparar los Writers Asíncronos
                var writerAsync = new ScheduleUpdateAsync();
                var viewActivator = new ViewActivatorAsync();

                // 8. Crear y mostrar ventana Modeless
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

        /// <summary>
        /// (Lógica sin cambios)
        /// </summary>
        private void RunDuplicateCheck(List<ElementData> elementosData)
        {
            // ... (Lógica de RunDuplicateCheck sin cambios) ...
            var metradosTablas = elementosData
                .Where(ed => ed.AuditResults.Any(ar => ar.AuditType == "FILTRO"))
                .ToList();
            var groupedByAC = metradosTablas.GroupBy(ed => ed.CodigoIdentificacion);
            foreach (var acGroup in groupedByAC)
            {
                if (acGroup.Count() < 2) continue;
                var groupedByType = acGroup.GroupBy(ed =>
                    (ed.Element as ViewSchedule)?.Definition.IsMaterialTakeoff ?? false
                );
                foreach (var typeGroup in groupedByType)
                {
                    if (typeGroup.Count() < 2) continue;
                    var groupedByCategory = typeGroup.GroupBy(ed =>
                        (ed.Element as ViewSchedule)?.Definition.CategoryId.ToString() ?? "NONE"
                    );
                    foreach (var categoryGroup in groupedByCategory)
                    {
                        if (categoryGroup.Count() < 2) continue;
                        foreach (var elementData in categoryGroup)
                        {
                            elementData.AuditResults.Add(new AuditItem
                            {
                                AuditType = "DUPLICADO",
                                Estado = EstadoParametro.Vacio,
                                Mensaje = "Advertencia: Tabla duplicada (mismo A.C., tipo y categoría).",
                                ValorActual = elementData.Nombre,
                                ValorCorregido = "N/A",
                                IsCorrectable = false
                            });
                            elementData.DatosCompletos = false;
                        }
                    }
                }
            }
        }
    }
}