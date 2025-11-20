// Plugins/Tablas/MainCommand.cs
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using TL60_AuditoriaUnificada.Models;
using TL60_AuditoriaUnificada.Core;
using TL60_AuditoriaUnificada.Plugins.Tablas;
using TL60_AuditoriaUnificada.UI;

namespace TL60_AuditoriaUnificada.Plugins.Tablas
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MainCommand : IExternalCommand
    {
        private const string SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";
        private const string HOJA_MODELOS_ESP = "MODELOS POR ESPECIALIDAD";

        private static readonly List<string> NOMBRES_WIP = new List<string>
        {
            "TL", "TITO", "PDONTADENEA", "ANDREA", "EFRAIN",
            "PROYECTOSBIM", "ASISTENTEBIM", "LUIS", "DIEGO", "JORGE", "MIGUEL",
            "AUDIT"
        };

        private static readonly Regex _acRegex = new Regex(@"^C\.(\d{2,3}\.)+\d{2,3}");

        private static readonly List<BuiltInCategory> _categoriesToAudit = new List<BuiltInCategory>
        {
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

                HashSet<string> modelWhitelist = GetModelWhitelistFromSheets(sheetsService, docTitle);

                var processor = new ScheduleProcessor(doc, sheetsService, uniclassService, SPREADSHEET_ID, mainSpecialty);

                var elementosData = new List<ElementData>();
                var existingMetradoCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                HashSet<string> codigosEncontradosEnModelos = new HashSet<string>();

                // Auditar Unidades Globales
                List<ElementData> unidadesGlobales = processor.AuditProjectUnits();

                // 2. PROCESAR TABLAS DE PLANIFICACIÓN
                var allSchedules = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Schedules)
                    .WhereElementIsNotElementType()
                    .Cast<ViewSchedule>()
                    .Where(v => v.IsTemplate == false);

                foreach (ViewSchedule view in allSchedules)
                {
                    if (view == null || view.Definition == null) continue;

                    string viewName = view.Name;
                    string viewNameUpper = viewName.ToUpper();
                    string grupoVista = GetParamValue(view, "GRUPO DE VISTA");
                    string subGrupoVista = GetParamValue(view, "SUBGRUPO DE VISTA");

                    if (viewNameUpper.StartsWith("SOPORTE."))
                    {
                        continue;
                    }

                    // ==========================================================
                    // RAMA 1: TABLAS DE METRADO (Empiezan con "C.")
                    // ==========================================================
                    if (viewNameUpper.StartsWith("C."))
                    {
                        if (viewNameUpper.Contains("COPY") || viewNameUpper.Contains("COPIA"))
                        {
                            if (grupoVista.Equals("REVISAR", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(subGrupoVista)) continue;
                            elementosData.Add(CreateReclassificationAudit(view,
                                "CLASIFICACIÓN (COPIA)", "Corregir: La tabla parece ser una copia.",
                                viewName.Replace("C.", "SOPORTE."), "REVISAR", "", null));
                            continue;
                        }

                        if (NOMBRES_WIP.Any(name => viewNameUpper.Contains(name)))
                        {
                            if (grupoVista.Equals("00 TRABAJO EN PROCESO - WIP", StringComparison.OrdinalIgnoreCase)) continue;
                            elementosData.Add(CreateReclassificationAudit(view,
                                "CLASIFICACIÓN (WIP)", "Corregir: Tabla de trabajo interno (WIP).",
                                viewName.Replace("C.", "SOPORTE."), "00 TRABAJO EN PROCESO - WIP", "SOPORTE BIM", "SOPORTE BIM"));
                            continue;
                        }

                        Match acMatch = _acRegex.Match(viewName);
                        string assemblyCode = acMatch.Success ? acMatch.Value : "INVALID_AC";

                        string scheduleType = uniclassService.GetScheduleType(assemblyCode);
                        if (scheduleType.Equals("MANUAL", StringComparison.OrdinalIgnoreCase))
                        {
                            if (grupoVista.Equals("00 TRABAJO EN PROCESO - WIP", StringComparison.OrdinalIgnoreCase) &&
                                subGrupoVista.Equals("METRADO MANUAL", StringComparison.OrdinalIgnoreCase)) continue;

                            elementosData.Add(CreateReclassificationAudit(view,
                                "MANUAL", "Corregir: Tabla de Metrado Manual.",
                                viewName.Replace("C.", "SOPORTE."), "00 TRABAJO EN PROCESO - WIP", "METRADO MANUAL", "SOPORTE CAMPO"));
                            continue;
                        }

                        if (grupoVista.StartsWith("C.", StringComparison.OrdinalIgnoreCase))
                        {
                            if (assemblyCode != "INVALID_AC")
                            {
                                existingMetradoCodes.Add(assemblyCode);
                            }

                            if (isModeEM) continue;

                            elementosData.Add(processor.ProcessSingleElement(view, assemblyCode));
                        }
                        else
                        {
                            elementosData.Add(CreateReclassificationAudit(view,
                                "CLASIFICACIÓN (SOPORTE)", "Corregir: Tabla de soporte mal clasificada.",
                                viewName.Replace("C.", "SOPORTE."), "00 TRABAJO EN PROCESO - WIP", "SOPORTE DE METRADOS", null));
                        }
                    }
                    // ==========================================================
                    // RAMA 2: OTRAS TABLAS (No empiezan con "C.")
                    // ==========================================================
                    else
                    {
                        var planos = GetPlanosDondeEstaLaTabla(doc, view);

                        // --- CASO 2A: TABLA ESTÁ "EN PLANO" ---
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
                        // ==========================================================
                        // --- CASO 2B: TABLA "NO PLANO" (LÓGICA ACTUALIZADA) ---
                        // ==========================================================
                        else
                        {
                            // 1. ¿ES CCECC? -> IGNORAR
                            if (grupoVista.Equals("AUDITORIA CCECC", StringComparison.OrdinalIgnoreCase))
                            {
                                // IGNORAR (No hacer nada)
                            }
                            // 2. ¿ES AUDIT? -> CLASIFICAR
                            else if (viewNameUpper.Contains("AUDIT"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (SOPORTE)", "Tabla de Auditoría.", null, "00 TRABAJO EN PROCESO - WIP", "SOPORTE BIM", "AUDITORÍA"));
                            }
                            // 3. ¿ES LOOK? -> CLASIFICAR
                            else if (viewNameUpper.Contains("LOOK"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (AVANCE)", "Tabla de Lookahead (No en plano).", null, "00 TRABAJO EN PROCESO - WIP", "AVANCE SEMANAL", "LOOKAHEAD"));
                            }
                            // 4. ¿ES AVANCE? -> CLASIFICAR
                            else if (viewNameUpper.Contains("AVANCE"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (AVANCE)", "Tabla de Avance (No en plano).", null, "00 TRABAJO EN PROCESO - WIP", "AVANCE SEMANAL", "CARS"));
                            }
                            // 5. ¿ES SECTORIZACI? -> CLASIFICAR
                            else if (viewNameUpper.Contains("SECTORIZACI"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (AVANCE)", "Tabla de Sectorización (No en plano).", null, "00 TRABAJO EN PROCESO - WIP", "AVANCE SEMANAL", "SECTORIZACIÓN"));
                            }
                            // 6. ¿ES VALORIZACI? -> CLASIFICAR
                            else if (viewNameUpper.Contains("VALORIZACI"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (SOPORTE)", "Tabla de Valorización (No en plano).", null, "00 TRABAJO EN PROCESO - WIP", "SOPORTE BIM", "VALORIZACIÓN"));
                            }
                            // 7. ¿ES COBIE? -> CLASIFICAR
                            else if (viewNameUpper.StartsWith("COBIE"))
                            {
                                elementosData.Add(CreateReclassificationAudit(view, "WIP (SOPORTE)", "Tabla COBie.", null, "00 TRABAJO EN PROCESO - WIP", "SOPORTE BIM", "COBIE"));
                            }
                            // 8. RESTO DE LAS TABLAS (CON REGLA DE PROTECCIÓN)
                            else
                            {
                                // Obtener el grupo de vista actual
                                string grupoActual = GetParamValue(view, "GRUPO DE VISTA");

                                // Aplicar la nueva regla de protección:
                                // Si ya está agrupada en una categoría principal, ignorarla.
                                if (grupoActual.Equals("00 TRABAJO EN PROCESO - WIP", StringComparison.OrdinalIgnoreCase) ||
                                    grupoActual.Equals("00 DISEÑO", StringComparison.OrdinalIgnoreCase) ||
                                    grupoActual.Equals("AUDITORIA CCECC", StringComparison.OrdinalIgnoreCase))
                                {
                                    // IGNORAR. Ya está clasificada, no hacer nada.
                                }
                                else
                                {
                                    // No está protegida, marcarla para REVISAR.
                                    elementosData.Add(CreateReclassificationAudit(view, "CLASIFICACIÓN (REVISAR)", "Tabla no clasificada.", null, "REVISAR", null, null));
                                }
                            }
                        }
                    }
                } // Fin del foreach

                // 4. AUDITORÍA DE ELEMENTOS (Tablas Faltantes)
                if (_categoriesToAudit.Count > 0)
                {
                    var elementAuditor = new MissingScheduleAuditor(doc, uniclassService);

                    ElementData missingSchedulesReport = elementAuditor.FindMissingSchedules(
                        existingMetradoCodes,
                        _categoriesToAudit,
                        modelWhitelist,
                        out codigosEncontradosEnModelos);

                    if (missingSchedulesReport != null)
                    {
                        elementosData.Add(missingSchedulesReport);
                    }
                }
                else
                {
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
                                Mensaje = "La auditoría de elementos modelados (Tablas Faltantes) no está configurada."
                            }
                        }
                    });
                }

                // 5. Auditoría de Tablas Metrado "Modo EM"
                if (isModeEM)
                {
                    foreach (ViewSchedule view in allSchedules)
                    {
                        if (!view.Name.StartsWith("C.") || !(GetParamValue(view, "GRUPO DE VISTA").StartsWith("C.")))
                            continue;

                        Match acMatch = _acRegex.Match(view.Name);
                        string assemblyCode = acMatch.Success ? acMatch.Value : "INVALID_AC";

                        if (!codigosEncontradosEnModelos.Contains(assemblyCode))
                        {
                            continue;
                        }

                        elementosData.Add(processor.ProcessSingleElement(view, assemblyCode));
                    }
                }

                // 6. POST-PROCESO: Auditoría de Duplicados
                RunDuplicateCheck(elementosData);

                // 7. Construir datos de diagnóstico
                var diagnosticBuilder = new DiagnosticDataBuilder();

                var todosElementos = new List<ElementData>();
                todosElementos.AddRange(unidadesGlobales);    // PRIMERO unidades
                todosElementos.AddRange(elementosData);       // DESPUÉS tablas

                var diagnosticRows = diagnosticBuilder.BuildDiagnosticRows(todosElementos);

                // 8. Preparar los Writers Asíncronos
                var writerAsync = new ScheduleUpdateAsync();
                var viewActivator = new ViewActivatorAsync();

                // 9. Crear y mostrar ventana Modeless
                var mainWindow = new MainWindow(
                    diagnosticRows,
                    todosElementos,  // Pasar la lista completa
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
        /// Helper para crear auditorías de reclasificación
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
                Categoria = "TABLAS", // Asignado a "TABLAS"
                CodigoIdentificacion = "N/A"
            };

            var jobData = new RenamingJobData
            {
                NuevoNombre = nuevoNombre,
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

            // Si todo ya está correcto, no crear auditoría
            if (nombreCorrecto && gvCorrecto && sgvCorrecto && ssvpCorrecto)
            {
                // No añadir a la lista. Devolver null para indicar que no hay nada que hacer.
                // Devolver un ElementData vacío pero "completo" podría llenar la UI
                // con filas "Correctas" innecesarias.
                // ... Reconsiderando: La lógica original devolvía un ElementData.
                // Lo mantendré, pero sin auditorías.
                elementData.DatosCompletos = true;
                return elementData; // Devuelve un elemento sin auditorías (no aparecerá en la UI)
            }

            elementData.DatosCompletos = false;
            var auditItem = new AuditItem
            {
                AuditType = auditType,
                IsCorrectable = true,
                Estado = EstadoParametro.Corregir,
                Mensaje = mensaje,
                ValorActual = $"Grupo: {gv}\nSubGrupo: {sgv}\nSubPartida: {ssvp}",
                ValorCorregido = $"Grupo: {nuevoGrupo ?? gv}\nSubGrupo: {nuevoSubGrupo ?? sgv}\nSubPartida: {nuevoSubPartida ?? ssvp}",
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

        #region Helpers de Clasificación

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

        private HashSet<string> GetModelWhitelistFromSheets(GoogleSheetsService sheetsService, string docTitle)
        {
            var whitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var data = sheetsService.ReadData(SPREADSHEET_ID, $"'{HOJA_MODELOS_ESP}'!B:B");
                if (data == null || data.Count == 0)
                {
                    whitelist.Add(docTitle);
                    return whitelist;
                }

                foreach (var row in data)
                {
                    if (row.Count == 0 || row[0] == null) continue;

                    string cellContent = row[0].ToString();
                    string[] modelsInCell = cellContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                    bool hostFoundInCell = false;
                    foreach (string modelName in modelsInCell)
                    {
                        if (Path.GetFileNameWithoutExtension(modelName.Trim()).Equals(docTitle, StringComparison.OrdinalIgnoreCase))
                        {
                            hostFoundInCell = true;
                            break;
                        }
                    }

                    if (hostFoundInCell)
                    {
                        foreach (string modelName in modelsInCell)
                        {
                            string trimmedName = Path.GetFileNameWithoutExtension(modelName.Trim());
                            if (!string.IsNullOrEmpty(trimmedName))
                            {
                                whitelist.Add(trimmedName);
                            }
                        }
                        return whitelist;
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error Google Sheets", $"No se pudo leer la hoja '{HOJA_MODELOS_ESP}': {ex.Message}");
            }

            whitelist.Add(docTitle);
            return whitelist;
        }

        #endregion
    }
}