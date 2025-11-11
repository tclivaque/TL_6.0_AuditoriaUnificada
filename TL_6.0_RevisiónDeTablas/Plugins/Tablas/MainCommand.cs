// Plugins/Tablas/MainCommand.cs
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using TL60_RevisionDeTablas.Models;
using TL60_RevisionDeTablas.Core;
using TL60_RevisionDeTablas.Plugins.Tablas;
using TL60_RevisionDeTablas.UI;
// using System.Text; // <--- Retirado

namespace TL60_RevisionDeTablas.Plugins.Tablas
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MainCommand : IExternalCommand
    {
        private const string SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";
        // private const string DEBUG_TARGET_AC = "..."; // <--- Retirado

        private static readonly List<string> NOMBRES_WIP = new List<string>
        {
            "TL", "TITO", "PDONTADENEA", "ANDREA", "EFRAIN",
            "PROYECTOSBIM", "ASISTENTEBIM", "LUIS", "DIEGO", "JORGE", "MIGUEL"
        };

        private static readonly Regex _acRegex = new Regex(@"^C\.(\d{2,3}\.)+\d{2,3}");

        private static readonly List<BuiltInCategory> _categoriesToAudit = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Roofs,
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_Ceilings,
            BuiltInCategory.OST_Stairs,
            BuiltInCategory.OST_StairsRailing,
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_EdgeSlab,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_ConduitFitting,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ElectricalEquipment,
            BuiltInCategory.OST_ElectricalFixtures,
            BuiltInCategory.OST_FlexPipeCurves,
            BuiltInCategory.OST_FlexDuctCurves,
            BuiltInCategory.OST_DataDevices,
            BuiltInCategory.OST_SecurityDevices,
            BuiltInCategory.OST_FireAlarmDevices,
            BuiltInCategory.OST_CommunicationDevices,
            BuiltInCategory.OST_NurseCallDevices,
            BuiltInCategory.OST_LightingFixtures,
            BuiltInCategory.OST_Furniture,
            BuiltInCategory.OST_FurnitureSystems,
            BuiltInCategory.OST_SpecialityEquipment
        };


        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // --- Debug Logger Retirado ---

            try
            {
                // 1. Inicializar Servicios
                var sheetsService = new GoogleSheetsService();
                string docTitle = Path.GetFileNameWithoutExtension(doc.Title);

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
                    if (view == null || view.Definition == null) continue;
                    string viewName = view.Name;
                    string viewNameUpper = viewName.ToUpper();
                    if (viewNameUpper.StartsWith("C."))
                    {
                        if (viewNameUpper.Contains("COPY") || viewNameUpper.Contains("COPIA"))
                        {
                            string grupoVista = view.LookupParameter("GRUPO DE VISTA")?.AsString() ?? string.Empty;
                            string subGrupoVista = view.LookupParameter("SUBGRUPO DE VISTA")?.AsString() ?? string.Empty;
                            if (grupoVista.Equals("REVISAR", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(subGrupoVista)) continue;
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
                        string scheduleType = uniclassService.GetScheduleType(assemblyCode);
                        if (scheduleType.Equals("MANUAL", StringComparison.OrdinalIgnoreCase))
                        {
                            string grupoVista = view.LookupParameter("GRUPO DE VISTA")?.AsString() ?? string.Empty;
                            string subGrupoVista = view.LookupParameter("SUBGRUPO DE VISTA")?.AsString() ?? string.Empty;
                            if (grupoVista.Equals("REVISAR", StringComparison.OrdinalIgnoreCase) && subGrupoVista.Equals("METRADO MANUAL", StringComparison.OrdinalIgnoreCase)) continue;
                            elementosData.Add(processor.CreateManualRenamingJob(view));
                            continue;
                        }
                        Parameter groupParam = view.LookupParameter("GRUPO DE VISTA");
                        string groupValue = groupParam?.AsString() ?? string.Empty;
                        if (groupValue.StartsWith("C.", StringComparison.OrdinalIgnoreCase))
                        {
                            elementosData.Add(processor.ProcessSingleElement(view, assemblyCode));
                            if (assemblyCode != "INVALID_AC")
                            {
                                existingMetradoCodes.Add(assemblyCode);
                            }
                        }
                        else
                        {
                            elementosData.Add(processor.CreateRenamingJob(view));
                        }
                    }
                }

                // --- Debug logs de Paso 2 retirados ---

                // ==========================================================
                // ===== 4. EJECUTAR AUDITORÍA DE ELEMENTOS (Tablas Faltantes)
                // ==========================================================

                if (_categoriesToAudit.Count > 0)
                {
                    var elementAuditor = new MissingScheduleAuditor(doc, uniclassService);

                    // (¡MODIFICADO!) Llamada limpia sin el logger 'sb'
                    ElementData missingSchedulesReport = elementAuditor.FindMissingSchedules(
                        existingMetradoCodes,
                        _categoriesToAudit);

                    if (missingSchedulesReport != null)
                    {
                        elementosData.Add(missingSchedulesReport);
                    }
                }
                else
                {
                    // ... (Advertencia de auditoría no configurada) ...
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
                                Estado = EstadoParametro.Vacio,
                                Mensaje = "La auditoría de elementos modelados (Tablas Faltantes) no está configurada. " +
                                          "Se debe definir la lista de categorías a revisar en MainCommand.cs."
                            }
                        }
                    });
                }

                // 5. POST-PROCESO: Auditoría de Duplicados
                RunDuplicateCheck(elementosData);

                // 6. Construir datos de diagnóstico
                var diagnosticBuilder = new DiagnosticDataBuilder();
                List<DiagnosticRow> diagnosticRows = diagnosticBuilder.BuildDiagnosticRows(elementosData);

                // --- Debug logs de Paso 5 retirados ---
                // --- Bloque de guardado de log retirado ---

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

                mainWindow.Show();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}\nStackTrace: {ex.StackTrace}";
                TaskDialog.Show("Error", message);
                // --- Bloque de guardado de log de error retirado ---
                return Result.Failed;
            }
        }

        /// <summary>
        /// Ejecuta la auditoría de duplicados
        /// </summary>
        private void RunDuplicateCheck(List<ElementData> elementosData)
        {
            // 1. Filtrar solo tablas de metrados (las que tienen auditorías de FILTRO)
            var metradosTablas = elementosData
                .Where(ed => ed.AuditResults.Any(ar => ar.AuditType == "FILTRO"))
                .ToList();

            // 2. Agrupar por Assembly Code
            var groupedByAC = metradosTablas.GroupBy(ed => ed.CodigoIdentificacion);

            foreach (var acGroup in groupedByAC)
            {
                if (acGroup.Count() < 2) continue; // No hay duplicados de A.C.

                // 3. Sub-agrupar por Tipo (Material vs Cantidades)
                var groupedByType = acGroup.GroupBy(ed =>
                    (ed.Element as ViewSchedule)?.Definition.IsMaterialTakeoff ?? false
                );

                foreach (var typeGroup in groupedByType)
                {
                    if (typeGroup.Count() < 2) continue;

                    // 4. Sub-agrupar por Categoría
                    var groupedByCategory = typeGroup.GroupBy(ed =>
                        (ed.Element as ViewSchedule)?.Definition.CategoryId.ToString() ?? "NONE"
                    );

                    foreach (var categoryGroup in groupedByCategory)
                    {
                        if (categoryGroup.Count() < 2) continue;

                        // 5. ¡Duplicado Encontrado! Marcar todos
                        foreach (var elementData in categoryGroup)
                        {
                            elementData.AuditResults.Add(new AuditItem
                            {
                                AuditType = "DUPLICADO",
                                Estado = EstadoParametro.Vacio, // Estado Advertencia
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