using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TL60_RevisionDeTablas.UI;
using TL60_RevisionDeTablas.Plugins.COBie.UI;
using TL60_RevisionDeTablas.Plugins.Tablas.UI;
using TL60_RevisionDeTablas.Models;
using TL60_RevisionDeTablas.Core;
using TL60_RevisionDeTablas.Plugins.COBie.Services;
using TL60_RevisionDeTablas.Plugins.COBie.Models;
using TL60_RevisionDeTablas.Plugins.Tablas;
using TL60_RevisionDeTablas.Plugins.Uniclass.Services;
using TL60_RevisionDeTablas.Plugins.Uniclass.UI;

namespace TL60_RevisionDeTablas.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class UnifiedAuditCommand : IExternalCommand
    {
        private const string COBIE_SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";
        private const string TABLAS_SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";

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

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // ====================================
                // PARTE 1: PROCESAR PLUGIN COBie
                // ====================================
                CobiePluginControl cobieControl = null;
                try
                {
                    cobieControl = ProcessCobiePlugin(doc);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error en COBie", $"Error al procesar plugin COBie:\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}");
                }

                // ====================================
                // PARTE 2: PROCESAR PLUGIN UNICLASS
                // ====================================
                UniclassPluginControl uniclassControl = null;
                try
                {
                    uniclassControl = ProcessUniclassPlugin(uidoc);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error en Uniclass", $"Error al procesar plugin Uniclass:\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}");
                }

                // ====================================
                // PARTE 3: PROCESAR PLUGIN TABLAS
                // ====================================
                TablasPluginControl tablasControl = null;
                try
                {
                    tablasControl = ProcessTablasPlugin(doc);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error en Tablas", $"Error al procesar plugin Tablas:\n{ex.Message}");
                }

                // ====================================
                // PARTE 4: MOSTRAR VENTANA UNIFICADA
                // ====================================
                if (cobieControl != null)
                {
                    var unifiedWindow = new UnifiedWindow(
                        metradosPlugin: null,          // Placeholder - no implementado aún
                        cobiePlugin: cobieControl,
                        uniclassPlugin: uniclassControl,
                        tablasPlugin: tablasControl,
                        projectBrowserPlugin: null     // Placeholder - no implementado aún
                    );
                    unifiedWindow.Show();
                    return Result.Succeeded;
                }
                else
                {
                    TaskDialog.Show("Error", "No se pudo cargar el plugin COBie.");
                    return Result.Failed;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error General", $"Ocurrió un error:\n{ex.Message}\n{ex.StackTrace}");
                message = ex.Message;
                return Result.Failed;
            }
        }

        private CobiePluginControl ProcessCobiePlugin(Document doc)
        {
            // Inicializar servicios
            var sheetsService = new GoogleSheetsService();
            sheetsService.Initialize();

            // Leer configuración
            var configManager = new ConfigurationManager(sheetsService);
            CobieConfig config = configManager.ReadConfiguration(COBIE_SPREADSHEET_ID);

            // Leer definiciones de parámetros
            var paramReader = new ParameterDefinitionReader(sheetsService);
            List<ParameterDefinition> definitions = paramReader.ReadParameterDefinitions(config.SpreadsheetId);

            // Leer matriz de control
            var controlMatrix = new ControlMatrixService(sheetsService);
            controlMatrix.LoadMatrix(config.SpreadsheetId);
            string docTitleWithoutExt = Path.GetFileNameWithoutExtension(doc.Title);
            ModelPermissions permissions = controlMatrix.GetPermissions(docTitleWithoutExt);

            // Procesar elementos
            var elementosData = new List<ElementData>();
            var facilityProcessor = new FacilityProcessor(doc, definitions);
            var floorProcessor = new FloorProcessor(doc, definitions, sheetsService, config);
            var roomProcessor = new RoomProcessor(doc, definitions, sheetsService, config);
            var elementProcessor = new ElementProcessor(doc, definitions, sheetsService, config);

            // Procesar según permisos (código simplificado del CobieCommand original)
            if (permissions.IsAllowed("FACILITY"))
            {
                var facilityData = facilityProcessor.ProcessFacility();
                if (facilityData != null) elementosData.Add(facilityData);
            }

            if (permissions.IsAllowed("FLOOR"))
            {
                var floorsData = floorProcessor.ProcessFloors();
                elementosData.AddRange(floorsData);
            }

            if (permissions.IsAllowed("SPACE"))
            {
                roomProcessor.Initialize();
                var roomsData = roomProcessor.ProcessRooms();
                elementosData.AddRange(roomsData);
            }

            bool typeAllowed = permissions.IsAllowed("TYPE");
            bool componentAllowed = permissions.IsAllowed("COMPONENT");

            if (typeAllowed || componentAllowed)
            {
                elementProcessor.Initialize();
                var elementsDataMep = elementProcessor.ProcessElements();
                if (!typeAllowed) elementsDataMep.RemoveAll(e => e.GrupoCOBie == "TYPE");
                if (!componentAllowed) elementsDataMep.RemoveAll(e => e.GrupoCOBie == "COMPONENT");
                elementosData.AddRange(elementsDataMep);
            }

            // Aplicar gráficos
            var viewManager = new ViewGraphicsManager(doc, doc.ActiveView, config.Categorias);
            if (viewManager.IsView3D())
            {
                viewManager.ApplyGraphicsAndIsolate(elementosData);
            }

            // Construir datos de diagnóstico (usar fully qualified name para evitar ambigüedad)
            var diagnosticBuilder = new TL60_RevisionDeTablas.Plugins.COBie.Services.DiagnosticDataBuilder();
            var diagnosticRows = diagnosticBuilder.BuildDiagnosticRows(elementosData);

            // Crear UserControl
            var cobieControl = new CobiePluginControl(
                diagnosticRows, elementosData, doc,
                facilityProcessor, floorProcessor, roomProcessor, elementProcessor, config.Categorias);

            return cobieControl;
        }

        private TablasPluginControl ProcessTablasPlugin(Document doc)
        {
            // Servicios
            var sheetsService = new GoogleSheetsService();
            string docTitle = Path.GetFileNameWithoutExtension(doc.Title);

            // Cargar Uniclass data
            var uniclassService = new UniclassDataService(sheetsService, docTitle);
            uniclassService.LoadClassificationData(TABLAS_SPREADSHEET_ID);

            // Crear procesador de tablas
            string mainSpecialty = GetSpecialtyFromTitle(docTitle);
            var processor = new ScheduleProcessor(doc, sheetsService, uniclassService, TABLAS_SPREADSHEET_ID, mainSpecialty);

            // Auditar Unidades Globales
            List<ElementData> unidadesGlobales = processor.AuditProjectUnits();

            // Procesar todas las tablas de planificación (lógica simplificada del MainCommand)
            var elementosData = new List<ElementData>();

            var allSchedules = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Schedules)
                .WhereElementIsNotElementType()
                .Cast<ViewSchedule>()
                .Where(v => v.IsTemplate == false);

            var existingMetradoCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Procesar cada tabla
            foreach (ViewSchedule view in allSchedules)
            {
                if (view == null || view.Definition == null) continue;

                string viewName = view.Name;
                if (viewName.ToUpper().StartsWith("SOPORTE.")) continue;

                // Si empieza con "C." y grupo de vista empieza con "C."
                if (viewName.StartsWith("C.") && GetParamValue(view, "GRUPO DE VISTA").StartsWith("C."))
                {
                    Match acMatch = _acRegex.Match(viewName);
                    string assemblyCode = acMatch.Success ? acMatch.Value : "INVALID_AC";

                    if (assemblyCode != "INVALID_AC")
                    {
                        existingMetradoCodes.Add(assemblyCode);
                    }

                    elementosData.Add(processor.ProcessSingleElement(view, assemblyCode));
                }
            }

            // Auditoría de elementos (tablas faltantes)
            var elementAuditor = new MissingScheduleAuditor(doc, uniclassService);
            HashSet<string> codigosEncontradosEnModelos;

            ElementData missingSchedulesReport = elementAuditor.FindMissingSchedules(
                existingMetradoCodes,
                _categoriesToAudit,
                new HashSet<string> { docTitle },
                out codigosEncontradosEnModelos);

            if (missingSchedulesReport != null)
            {
                elementosData.Add(missingSchedulesReport);
            }

            // Construir datos de diagnóstico
            var diagnosticBuilder = new TL60_RevisionDeTablas.Plugins.Tablas.DiagnosticDataBuilder();

            var todosElementos = new List<ElementData>();
            todosElementos.AddRange(unidadesGlobales);
            todosElementos.AddRange(elementosData);

            var diagnosticRows = diagnosticBuilder.BuildDiagnosticRows(todosElementos);

            // Preparar writers asíncronos
            var writerAsync = new ScheduleUpdateAsync();
            var viewActivator = new ViewActivatorAsync();

            // Crear y retornar UserControl
            var tablasControl = new TablasPluginControl(
                diagnosticRows,
                todosElementos,
                doc,
                writerAsync,
                viewActivator
            );

            return tablasControl;
        }

        private UniclassPluginControl ProcessUniclassPlugin(UIDocument uidoc)
        {
            Document doc = uidoc.Document;

            // Inicializar servicios
            var sheetsService = new GoogleSheetsService();
            sheetsService.Initialize();

            // Inicializar Uniclass data service
            string docTitleWithoutExt = Path.GetFileNameWithoutExtension(doc.Title);
            var uniclassService = new UniclassDataService(sheetsService, docTitleWithoutExt);
            uniclassService.LoadClassificationData(TABLAS_SPREADSHEET_ID);

            // Leer categorías a auditar desde Google Sheets
            List<BuiltInCategory> categoriasAuditar = ReadCategoriasFromSheet(sheetsService, TABLAS_SPREADSHEET_ID);

            if (categoriasAuditar.Count == 0)
            {
                // Si no se encuentran categorías, retornar null (no mostrar pestaña)
                System.Diagnostics.Debug.WriteLine("WARN: No se encontraron categorías para Uniclass");
                return null;
            }

            // Procesar elementos
            var auditService = new UniclassAuditService(doc, uniclassService, categoriasAuditar);
            var elementosData = auditService.ProcessElements();

            if (elementosData.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("INFO: No se encontraron elementos para auditar en Uniclass");
                return null;
            }

            // Crear UserControl
            var uniclassControl = new UniclassPluginControl(uidoc, elementosData);

            return uniclassControl;
        }

        #region Helpers

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

        private string GetSpecialtyFromTitle(string docTitle)
        {
            try
            {
                var parts = Path.GetFileNameWithoutExtension(docTitle).Split('-');
                if (parts.Length > 3)
                {
                    return parts[3];
                }
            }
            catch (Exception) { }
            return "UNKNOWN_SPECIALTY";
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

        #endregion
    }
}
