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

namespace TL60_RevisionDeTablas.Plugins.Tablas
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MainCommand : IExternalCommand
    {
        private const string SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";

        // (¡NUEVO!) Lista de WIP expandida, basada en TL_5.0
        private static readonly List<string> NOMBRES_WIP = new List<string>
        {
            "TL", "TITO", "PDONTADENEA", "ANDREA", "EFRAIN",
            "PROYECTOSBIM", "ASISTENTEBIM", "LUIS", "DIEGO", "JORGE", "MIGUEL",
            "LOOKAHEAD", "CARS", "SECTORIZACIÓN", "VALORIZACIÓN", "AUDIT"
        };

        // (¡NUEVO!) Lista blanca para Modo Restringido EM
        private static readonly List<string> EM_WHITELIST = new List<string>
        {
            "200114-CCC02-MO-EM-000410",
            "200114-CCC02-MO-EM-045500",
            "200114-CCC02-MO-EM-045600"
        };

        private static readonly Regex _acRegex = new Regex(@"^C\.(\d{2,3}\.)+\d{2,3}");

        private static readonly List<BuiltInCategory> _categoriesToAudit = new List<BuiltInCategory>
        {
            // ... (Lista de 31 categorías sin cambios) ...
            BuiltInCategory.OST_StructuralColumns, BuiltInCategory.OST_StructuralFraming, BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Roofs, BuiltInCategory.OST_Walls, BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_Ceilings, BuiltInCategory.OST_Stairs, BuiltInCategory.OST_StairsRailing,
            BuiltInCategory.OST_Doors, BuiltInCategory.OST_Windows, BuiltInCategory.OST_EdgeSlab,
            BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ElectricalEquipment, BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_FlexPipeCurves,
            BuiltInCategory.OST_FlexDuctCurves, BuiltInCategory.OST_DataDevices, BuiltInCategory.OST_SecurityDevices,
            BuiltInCategory.OST_FireAlarmDevices, BuiltInCategory.OST_CommunicationDevices, BuiltInCategory.OST_NurseCallDevices,
            BuiltInCategory.OST_LightingFixtures, BuiltInCategory.OST_Furniture, BuiltInCategory.OST_FurnitureSystems,
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

            try
            {
                // 1. Inicializar Servicios
                var sheetsService = new GoogleSheetsService();
                string docTitle = Path.GetFileNameWithoutExtension(doc.Title);
                string mainSpecialty = GetSpecialtyFromTitle(docTitle);
                bool isModeEM = mainSpecialty.Equals("EM", StringComparison.OrdinalIgnoreCase);

                var uniclassService = new UniclassDataService(sheetsService, docTitle);
                uniclassService.LoadClassificationData(SPREADSHEET_ID);

                // (¡MODIFICADO!) Pasamos la especialidad al procesador
                var processor = new ScheduleProcessor(doc, sheetsService, uniclassService, SPREADSHEET_ID, mainSpecialty);

                var elementosData = new List<ElementData>();
                var existingMetradoCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                HashSet<string> codigosEmRelevantes = new HashSet<string>();

                // ==========================================================
                // ===== 2. EJECUTAR AUDITORÍA DE ELEMENTOS (Tablas Faltantes)
                // ==========================================================
                // (¡MODIFICADO!) Se ejecuta antes si estamos en Modo EM
                // ==========================================================

                if (_categoriesToAudit.Count > 0)
                {
                    var elementAuditor = new MissingScheduleAuditor(doc, uniclassService);

                    // (¡MODIFICADO!) La firma del método cambió
                    ElementData missingSchedulesReport = elementAuditor.FindMissingSchedules(
                        existingMetradoCodes, // Pasa lista vacía (aún no se llena)
                        _categoriesToAudit,
                        out codigosEmRelevantes); // <--- (¡NUEVO!) Obtiene los códigos

                    if (missingSchedulesReport != null)
                    {
                        elementosData.Add(missingSchedulesReport);
                    }

                    // Si no es EM, 'codigosEmRelevantes' contendrá todos los AC de la especialidad.
                    // Si es EM, solo contendrá los AC de la "lista blanca".
                }
                else
                {
                    // ... (Advertencia de auditoría no configurada) ...
                }

                // ==========================================================
                // ===== 3. PROCESAR TABLAS DE PLANIFICACIÓN (ÁRBOL DE DECISIÓN)
                // ==========================================================

                var allSchedules = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Schedules)
                    .WhereElementIsNotElementType()
                    .Cast<ViewSchedule>();

                foreach (ViewSchedule view in allSchedules)
                {
                    if (view == null || view.Definition == null) continue;

                    string viewName = view.Name;
                    string viewNameUpper = viewName.ToUpper();
                    string grupoVista = GetParamValue(view, "GRUPO DE VISTA");
                    string subGrupoVista = GetParamValue(view, "SUBGRUPO DE VISTA");

                    // ==========================================================
                    // ===== INICIO: RAMA SUPERIOR (tabla.Name empieza con "C.")
                    // ==========================================================
                    if (viewNameUpper.StartsWith("C."))
                    {
                        // --- PREGUNTA 2: ¿ES COPIA? ---
                        if (viewNameUpper.Contains("COPY") || viewNameUpper.Contains("COPIA"))
                        {
                            // Evitar falso positivo si ya está corregido
                            if (grupoVista.Equals("REVISAR", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(subGrupoVista)) continue;

                            elementosData.Add(CreateReclassificationAudit(view,
                                "CLASIFICACIÓN (COPIA)",
                                "Corregir: La tabla parece ser una copia.",
                                viewName.Replace("C.", "SOPORTE."),
                                "REVISAR",
                                "",
                                null));
                            continue;
                        }

                        // --- PREGUNTA 3: ¿ES WIP? ---
                        // (¡MODIFICADO!) Usamos la nueva lista de WIP
                        if (NOMBRES_WIP.Any(name => viewNameUpper.Contains(name)))
                        {
                            elementosData.Add(CreateReclassificationAudit(view,
                                "CLASIFICACIÓN (WIP)",
                                "Corregir: Tabla de trabajo interno (WIP).",
                                viewName.Replace("C.", "SOPORTE."),
                                "00 TRABAJO EN PROCESO - WIP",
                                "SOPORTE BIM",
                                null)); // Asumimos destino "SOPORTE BIM" para nombres de personas
                            continue;
                        }

                        Match acMatch = _acRegex.Match(viewName);
                        string assemblyCode = acMatch.Success ? acMatch.Value : "INVALID_AC";

                        // --- PREGUNTA 4: ¿ES MANUAL? ---
                        string scheduleType = uniclassService.GetScheduleType(assemblyCode);
                        if (scheduleType.Equals("MANUAL", StringComparison.OrdinalIgnoreCase))
                        {
                            // Evitar falso positivo si ya está corregido
                            if (grupoVista.Equals("00 TRABAJO EN PROCESO - WIP", StringComparison.OrdinalIgnoreCase) &&
                                subGrupoVista.Equals("METRADO MANUAL", StringComparison.OrdinalIgnoreCase)) continue;

                            elementosData.Add(CreateReclassificationAudit(view,
                                "MANUAL",
                                "Corregir: Tabla de Metrado Manual.",
                                viewName.Replace("C.", "SOPORTE."),
                                "00 TRABAJO EN PROCESO - WIP",
                                "METRADO MANUAL",
                                "SOPORTE CAMPO")); // <-- (¡NUEVO!) Subpartida
                            continue;
                        }

                        // --- PREGUNTA 5: ¿ES TABLA DE METRADO? ---
                        if (grupoVista.StartsWith("C.", StringComparison.OrdinalIgnoreCase))
                        {
                            // (¡NUEVO!) Aplicar filtro de "Modo EM"
                            if (isModeEM && !codigosEmRelevantes.Contains(assemblyCode))
                            {
                                // Si es modo EM y el AC de esta tabla no se
                                // encontró en los modelos de la whitelist, ignorarla.
                                continue;
                            }

                            // Es una tabla de metrado válida, auditarla a fondo
                            elementosData.Add(processor.ProcessSingleElement(view, assemblyCode));
                            if (assemblyCode != "INVALID_AC")
                            {
                                existingMetradoCodes.Add(assemblyCode);
                            }
                        }
                        // --- ES TABLA DE SOPORTE ---
                        else
                        {
                            elementosData.Add(CreateReclassificationAudit(view,
                                "CLASIFICACIÓN (SOPORTE)",
                                "Corregir: Tabla de soporte mal clasificada.",
                                viewName.Replace("C.", "SOPORTE."),
                                "00 TRABAJO EN PROCESO - WIP",
                                "SOPORTE DE METRADOS",
                                null));
                        }
                    }
                    // ==========================================================
                    // ===== INICIO: RAMA INFERIOR (tabla.Name NO empieza con "C.")
                    // ==========================================================
                    else
                    {
                        var planos = GetPlanosDondeEstaLaTabla(doc, view);

                        // --- PREGUNTA 6: ¿ESTÁ EN PLANO? ---
                        if (planos.Count > 0)
                        {
                            ViewSheet primerPlano = planos[0];
                            string nombrePlano = primerPlano.Name.ToUpper();

                            if (EstaEnPlanoDeEntrega(primerPlano))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (DISEÑO)", "Tabla en plano de Entrega.", null, "00 DISEÑO", null, null));
                            }
                            else if (nombrePlano.Contains("LOOK"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (AVANCE)", "Tabla en plano Lookahead.", null, "00 TRABAJO EN PROCESO - WIP", "AVANCE SEMANAL", "LOOKAHEAD"));
                            }
                            else if (nombrePlano.Contains("AVANCE"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (AVANCE)", "Tabla en plano de Avance.", null, "00 TRABAJO EN PROCESO - WIP", "AVANCE SEMANAL", "CARS"));
                            }
                            else if (nombrePlano.Contains("SECTORIZACI"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (AVANCE)", "Tabla en plano de Sectorización.", null, "00 TRABAJO EN PROCESO - WIP", "AVANCE SEMANAL", "SECTORIZACIÓN"));
                            }
                            else if (nombrePlano.Contains("VALORIZACI"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (SOPORTE)", "Tabla en plano de Valorización.", null, "00 TRABAJO EN PROCESO - WIP", "SOPORTE BIM", "VALORIZACIÓN"));
                            }
                            else // Default si está en un plano no reconocido
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (SOPORTE)", "Tabla en plano de Soporte Campo.", null, "00 TRABAJO EN PROCESO - WIP", "SOPORTE BIM", "SOPORTE CAMPO"));
                            }
                        }
                        // --- NO ESTÁ EN PLANO ---
                        else
                        {
                            // --- PREGUNTA 7: ¿ES CCECC? ---
                            if (grupoVista.Equals("AUDITORIA CCECC", StringComparison.OrdinalIgnoreCase))
                            {
                                // IGNORAR (No hacer nada)
                            }
                            // --- PREGUNTA 8: ¿ES AUDIT? ---
                            else if (viewNameUpper.Contains("AUDIT"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (SOPORTE)", "Tabla de Auditoría.", null, "00 TRABAJO EN PROCESO - WIP", "SOPORTE BIM", "AUDITORÍA"));
                            }
                            // --- PREGUNTA 9: ¿ES COBIE? ---
                            else if (viewNameUpper.StartsWith("COBIE"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (SOPORTE)", "Tabla COBie.", null, "00 TRABAJO EN PROCESO - WIP", "SOPORTE BIM", "COBIE")); // Asumimos subpartida COBIE
                            }
                            // --- PREGUNTA 10: RESTO ---
                            else
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "CLASIFICACIÓN (REVISAR)", "Tabla no clasificada.", null, "REVISAR", null, null));
                            }
                        }
                    }
                } // Fin del foreach

                // 5. POST-PROCESO: Auditoría de Duplicados
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

                mainWindow.Show();
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
        /// (¡NUEVO!) Helper para crear auditorías de reclasificación directamente
        /// </summary>
        private ElementData CreateReclassificationAudit(
            ViewSchedule view, string auditType, string mensaje,
            string nuevoNombre, string nuevoGrupo, string nuevoSubGrupo, string nuevoSubPartida)
        {
            var elementData = new ElementData
            {
                ElementId = view.Id,
                Element = view,
                Nombre = view.Name,
                Categoria = view.Category?.Name ?? "Tabla de Planificación",
                CodigoIdentificacion = "N/A"
            };

            var jobData = new RenamingJobData
            {
                NuevoNombre = nuevoNombre, // Si es null, el writer lo ignorará
                NuevoGrupoVista = nuevoGrupo,
                NuevoSubGrupoVista = nuevoSubGrupo,
                NuevoSubGrupoVistaSubpartida = nuevoSubPartida
            };

            // Comprobar si realmente hay algo que corregir
            string gv = GetParamValue(view, "GRUPO DE VISTA");
            string sgv = GetParamValue(view, "SUBGRUPO DE VISTA");
            string ssvp = GetParamValue(view, "SUBGRUPO DE VISTA_SUBPARTIDA");

            bool nombreCorrecto = (nuevoNombre == null) || (view.Name == nuevoNombre);
            bool gvCorrecto = (nuevoGrupo == null) || (gv == nuevoGrupo);
            bool sgvCorrecto = (nuevoSubGrupo == null) || (sgv == nuevoSubGrupo);
            bool ssvpCorrecto = (nuevoSubPartida == null) || (ssvp == nuevoSubPartida);

            if (nombreCorrecto && gvCorrecto && sgvCorrecto && ssvpCorrecto)
            {
                elementData.DatosCompletos = true; // Ya está todo correcto
                return elementData; // No añadir auditoría
            }

            // Si algo está mal, crear la auditoría
            elementData.DatosCompletos = false;
            var auditItem = new AuditItem
            {
                AuditType = auditType,
                IsCorrectable = true,
                Estado = EstadoParametro.Corregir,
                Mensaje = mensaje,
                ValorActual = $"Grupo: {gv}\nSubGrupo: {sgv}\nSubPartida: {ssvp}",
                ValorCorregido = $"Grupo: {nuevoGrupo}\nSubGrupo: {nuevoSubGrupo}\nSubPartida: {nuevoSubPartida}",
                Tag = jobData
            };

            if (nuevoNombre != null)
            {
                auditItem.ValorActual = $"Nombre: {view.Name}\n" + auditItem.ValorActual;
                auditItem.ValorCorregido = $"Nombre: {nuevoNombre}\n" + auditItem.ValorCorregido;
            }

            elementData.AuditResults.Add(auditItem);
            return elementData;
        }

        // ... (RunDuplicateCheck sin cambios) ...
        #region Helpers (Sin Cambios)
        private void RunDuplicateCheck(List<ElementData> elementosData)
        {
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
        #endregion

        // ==========================================================
        // ===== (¡NUEVO!) Helpers importados de TL_3_OrganizarTablas.cs
        // ==========================================================
        #region Helpers de Clasificación (TL_5.0)

        private string GetParamValue(Element elemento, string nombre_param)
        {
            if (elemento == null) return string.Empty;
            Parameter param = elemento.LookupParameter(nombre_param);
            if (param != null && param.HasValue)
            {
                return param.AsString() ?? string.Empty;
            }
            return string.Empty;
        }

        private List<ViewSheet> GetPlanosDondeEstaLaTabla(Document doc, ViewSchedule tabla)
        {
            var scheduleInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(ScheduleSheetInstance))
                .Cast<ScheduleSheetInstance>()
                .Where(inst => inst.ScheduleId == tabla.Id);

            var planos = new List<ViewSheet>();
            foreach (ScheduleSheetInstance inst in scheduleInstances)
            {
                ViewSheet plano = doc.GetElement(inst.OwnerViewId) as ViewSheet;
                if (plano != null)
                {
                    planos.Add(plano);
                }
            }
            return planos.Distinct().ToList();
        }

        private bool EstaEnPlanoDeEntrega(ViewSheet plano)
        {
            if (plano == null) return false;
            try
            {
                string grupo_plano = GetParamValue(plano, "GRUPO DE VISTA").ToUpper();
                string subgrupo_plano = GetParamValue(plano, "SUBGRUPO DE VISTA").ToUpper();
                if (grupo_plano.Contains("ENTREGA") || subgrupo_plano.Contains("ENTREGA"))
                {
                    return true;
                }
                return false;
            }
            catch (Exception) { return false; }
        }

        private string GetSpecialtyFromTitle(string docTitle)
        {
            try
            {
                var parts = Path.GetFileNameWithoutExtension(docTitle).Split('-');
                if (parts.Length > 3)
                {
                    return parts[3]; // "ES", "EE", "AR", etc.
                }
            }
            catch (Exception) { /* Ignorar */ }
            return "UNKNOWN_SPECIALTY";
        }

        #endregion
    }
}