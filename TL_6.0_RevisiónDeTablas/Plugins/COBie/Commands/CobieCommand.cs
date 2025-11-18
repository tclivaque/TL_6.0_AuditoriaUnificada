using System;
using TL60_RevisionDeTablas.Core;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture; // Necesario para Room
using Autodesk.Revit.DB.Mechanical; // Necesario para MEPCurve
using Autodesk.Revit.DB.Plumbing; // Necesario para Pipe
using Autodesk.Revit.DB.Electrical; // Necesario para Conduit
using Autodesk.Revit.UI;
using TL60_RevisionDeTablas.Plugins.COBie.Models;
using TL60_RevisionDeTablas.Models;
using TL60_RevisionDeTablas.Plugins.COBie.Services;
using TL60_RevisionDeTablas.Plugins.COBie.UI;
using System.Globalization;


namespace TL60_RevisionDeTablas.Plugins.COBie.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CobieCommand : IExternalCommand
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

                // PASO 2: LEER CONFIGURACIÓN
                var configManager = new ConfigurationManager(sheetsService);
                CobieConfig config = configManager.ReadConfiguration(SPREADSHEET_ID);

                // PASO 3: LEER DEFINICIONES DE PARÁMETROS
                var paramReader = new ParameterDefinitionReader(sheetsService);
                List<ParameterDefinition> definitions = paramReader.ReadParameterDefinitions(config.SpreadsheetId);

                // PASO 4: LEER MATRIZ DE CONTROL Y OBTENER PERMISOS
                var controlMatrix = new ControlMatrixService(sheetsService);
                controlMatrix.LoadMatrix(config.SpreadsheetId);
                string docTitleWithoutExt = Path.GetFileNameWithoutExtension(doc.Title);
                ModelPermissions permissions = controlMatrix.GetPermissions(docTitleWithoutExt);

                // PASO 5: INICIALIZAR PROCESADORES Y LISTA DE DATOS
                var elementosData = new List<ElementData>();
                var facilityProcessor = new FacilityProcessor(doc, definitions);
                var floorProcessor = new FloorProcessor(doc, definitions, sheetsService, config);
                var roomProcessor = new RoomProcessor(doc, definitions, sheetsService, config);
                var elementProcessor = new ElementProcessor(doc, definitions, sheetsService, config);
                // El ParameterWriter se instancia solo si es necesario (dentro del Handler)

                // PASO 6: PROCESAR ELEMENTOS O GENERAR DATOS POR DEFECTO
                // --- FACILITY ---
                if (permissions.IsAllowed("FACILITY"))
                {
                    var facilityData = facilityProcessor.ProcessFacility();
                    if (facilityData != null) elementosData.Add(facilityData);
                }
                else
                {
                    elementosData.AddRange(GenerateDefaultDataForGroup("FACILITY", config, definitions, doc));
                }

                // --- FLOOR ---
                if (permissions.IsAllowed("FLOOR"))
                {
                    var floorsData = floorProcessor.ProcessFloors();
                    elementosData.AddRange(floorsData);
                }
                else
                {
                    elementosData.AddRange(GenerateDefaultDataForGroup("FLOOR", config, definitions, doc));
                }

                // --- SPACE ---
                if (permissions.IsAllowed("SPACE"))
                {
                    roomProcessor.Initialize();
                    var roomsData = roomProcessor.ProcessRooms();
                    elementosData.AddRange(roomsData);
                }
                else
                {
                    elementosData.AddRange(GenerateDefaultDataForGroup("SPACE", config, definitions, doc));
                }

                // --- TYPE & COMPONENT ---
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

                if (!typeAllowed)
                {
                    foreach (var catStr in config.Categorias)
                    {
                        if (Enum.TryParse(catStr, out BuiltInCategory bic))
                        {
                            elementosData.AddRange(GenerateDefaultDataForGroup("TYPE", config, definitions, doc, bic));
                        }
                    }
                }
                if (!componentAllowed)
                {
                    foreach (var catStr in config.Categorias)
                    {
                        if (Enum.TryParse(catStr, out BuiltInCategory bic))
                        {
                            elementosData.AddRange(GenerateDefaultDataForGroup("COMPONENT", config, definitions, doc, bic));
                        }
                    }
                }

                // PASO 7: APLICAR GRÁFICOS
                var viewManager = new ViewGraphicsManager(doc, doc.ActiveView, config.Categorias);
                if (viewManager.IsView3D())
                {
                    viewManager.ApplyGraphicsAndIsolate(elementosData);
                }

                // PASO 8: CONSTRUIR DATOS DIAGNÓSTICO
                var diagnosticBuilder = new DiagnosticDataBuilder();
                var diagnosticRows = diagnosticBuilder.BuildDiagnosticRows(elementosData);

                // PASO 9: MOSTRAR VENTANA
                var diagnosticWindow = new DiagnosticWindow(
                    diagnosticRows, elementosData, doc,
                    facilityProcessor, floorProcessor, roomProcessor, elementProcessor, config.Categorias);
                diagnosticWindow.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error en COBie Automation", $"Ocurrió un error:\n{ex.Message}\n{ex.StackTrace}");
                message = ex.Message;
                return Result.Failed;
            }
        }

        /// <summary>
        /// Genera ElementData con valores por defecto para grupos N/A.
        /// </summary>
        private List<ElementData> GenerateDefaultDataForGroup(string grupo, CobieConfig config, List<ParameterDefinition> definitions, Document doc, BuiltInCategory? categoria = null)
        {
            var defaultElementDataList = new List<ElementData>();
            var elements = GetElementsByGroup(grupo, doc, categoria);
            if (!elements.Any()) return defaultElementDataList;

            var paramDefinitions = definitions
                .Where(d => d.Grupo.Equals(grupo, StringComparison.OrdinalIgnoreCase)
                         || d.Grupo.IndexOf($". {grupo}", StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            foreach (var elem in elements)
            {
                var ed = new ElementData
                {
                    Element = elem,
                    ElementId = elem.Id,
                    Nombre = elem.Name,
                    GrupoCOBie = grupo,
                    Categoria = elem.Category?.Name ?? "N/A",
                    CobieCompleto = false // Se recalculará después
                };

                Element targetElement = elem;
                if (grupo == "TYPE")
                {
                    ElementId typeId = GetElementTypeIdHelper(elem);
                    targetElement = (typeId != null && typeId != ElementId.InvalidElementId) ? doc.GetElement(typeId) : null;
                }

                if (targetElement == null) continue;

                var processedParams = new HashSet<string>(); // Para evitar duplicados en método híbrido

                // Método 1: Usar definiciones explícitas
                foreach (var def in paramDefinitions)
                {
                    Parameter param = targetElement.LookupParameter(def.NombreParametro);
                    if (param != null)
                    {
                        // ===== INICIO DE CORRECCIÓN =====
                        string valorActual = GetParameterValueHelper(param);
                        string defaultValue = GetDefaultValueForParameterTypeHelper(param, config);

                        if (AreValuesEqualHelper(valorActual, defaultValue))
                        {
                            // Ya está correcto
                            ed.ParametrosCorrectos[def.NombreParametro] = valorActual;
                        }
                        else
                        {
                            // Necesita corrección
                            ed.ParametrosActualizar[def.NombreParametro] = defaultValue;
                            if (string.IsNullOrWhiteSpace(valorActual))
                            {
                                ed.ParametrosVacios.Add(def.NombreParametro);
                            }
                        }
                        // ===== FIN DE CORRECCIÓN =====

                        processedParams.Add(def.NombreParametro); // Marcar como procesado
                    }
                }

                // Método 2: Buscar "COBie.*" restantes (solo TYPE/COMPONENT)
                if (grupo == "TYPE" || grupo == "COMPONENT")
                {
                    foreach (Parameter param in targetElement.Parameters)
                    {
                        if (param == null) continue;
                        string paramName = param.Definition.Name;
                        if (paramName.StartsWith("COBie", StringComparison.OrdinalIgnoreCase) && !processedParams.Contains(paramName))
                        {
                            // ===== INICIO DE CORRECCIÓN (REPETIDA) =====
                            string valorActual = GetParameterValueHelper(param);
                            string defaultValue = GetDefaultValueForParameterTypeHelper(param, config);

                            if (AreValuesEqualHelper(valorActual, defaultValue))
                            {
                                // Ya está correcto
                                ed.ParametrosCorrectos[paramName] = valorActual;
                            }
                            else
                            {
                                // Necesita corrección
                                ed.ParametrosActualizar[paramName] = defaultValue;
                                if (string.IsNullOrWhiteSpace(valorActual))
                                {
                                    ed.ParametrosVacios.Add(paramName);
                                }
                            }
                            // ===== FIN DE CORRECCIÓN (REPETIDA) =====
                        }
                    }
                }

                // Calcular si COBie está completo
                ed.CobieCompleto = ed.ParametrosVacios.Count == 0 &&
                                   ed.ParametrosActualizar.Count == 0 && // <-- Añadido
                                   ed.Mensajes.Count == 0;

                defaultElementDataList.Add(ed);
            }
            return defaultElementDataList;
        }

        // ===== CÓDIGO INACCESIBLE ELIMINADO: Los métodos helper ya no están aquí =====

        /// <summary>
        /// Helper para obtener elementos según el grupo COBie (Necesario aquí)
        /// </summary>
        private List<Element> GetElementsByGroup(string grupo, Document doc, BuiltInCategory? categoria = null)
        {
            var elements = new List<Element>();
            switch (grupo)
            {
                case "FACILITY":
                    elements.AddRange(new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectInformation).ToElements());
                    break;
                case "FLOOR":
                    elements.AddRange(new FilteredElementCollector(doc).OfClass(typeof(Level)).WhereElementIsNotElementType().ToElements());
                    break;
                case "SPACE":
                    elements.AddRange(new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
                                      .Cast<Room>().Where(r => r.Area > 0).Cast<Element>());
                    break;
                case "TYPE":
                case "COMPONENT":
                    if (categoria.HasValue) elements.AddRange(new FilteredElementCollector(doc).OfCategory(categoria.Value).WhereElementIsNotElementType().ToElements());
                    break;
            }
            return elements;
        }

        /// <summary>
        /// Helper para obtener el valor por defecto según el tipo de parámetro (Necesario aquí)
        /// </summary>
        private string GetDefaultValueForParameterTypeHelper(Parameter param, CobieConfig config)
        {
            if (param == null) return config.ValorDefectoString;
            switch (param.StorageType)
            {
                case StorageType.Integer:
                    return (param.Definition.ParameterType == ParameterType.YesNo) ? config.ValorDefectoIntegerYesNo : config.ValorDefectoInteger;
                case StorageType.Double:
                    return config.ValorDefectoDouble;
                case StorageType.String:
                default:
                    return config.ValorDefectoString;
            }
        }

        /// <summary>
        /// Helper para obtener el ElementTypeId (copiado de ParameterWriter)
        /// </summary>
        private ElementId GetElementTypeIdHelper(Element elemento)
        {
            if (elemento == null) return null;
            if (elemento is FamilyInstance fi) return fi.GetTypeId();
            if (elemento is Wall wall) return wall.GetTypeId();
            if (elemento is Floor floor) return floor.GetTypeId();
            if (elemento is Ceiling ceiling) return ceiling.GetTypeId();
            if (elemento is RoofBase roof) return roof.GetTypeId();
            if (elemento is Stairs stairs) return stairs.GetTypeId();
            if (elemento is Railing railing) return railing.GetTypeId();
            if (elemento is MEPCurve mepCurve) return mepCurve.GetTypeId();
            try { return elemento.GetTypeId(); } catch { return null; }
            // ===== CAMBIO AQUÍ: Eliminado el return null final =====
        }

        /// <summary>
        /// Helper para obtener valor de parámetro (copiado de processors)
        /// </summary>
        private string GetParameterValueHelper(Parameter param)
        {
            if (param == null || !param.HasValue) return string.Empty;
            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        return param.AsString() ?? string.Empty;
                    case StorageType.Integer:
                        return param.AsInteger().ToString();
                    case StorageType.Double:
                        if (!param.HasValue)
                        {
                            return string.Empty;
                        }
                        double valorInterno = param.AsDouble();
                        ForgeTypeId unitTypeId = param.GetUnitTypeId();
                        if (unitTypeId != null &&
                            (unitTypeId.Equals(UnitTypeId.SquareFeet) ||
                             unitTypeId.Equals(UnitTypeId.SquareMeters)))
                        {
                            double valorConvertido = UnitUtils.ConvertFromInternalUnits(valorInterno, UnitTypeId.SquareMeters);
                            return valorConvertido.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            double valorConvertido = UnitUtils.ConvertFromInternalUnits(valorInterno, UnitTypeId.Meters);
                            return valorConvertido.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                        }
                    default:
                        return param.AsValueString() ?? string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Helper para comparar valores (copiado de processors)
        /// </summary>
        private bool AreValuesEqualHelper(string value1, string value2)
        {
            string val1 = string.IsNullOrWhiteSpace(value1) ? string.Empty : value1.Trim();
            string val2 = string.IsNullOrWhiteSpace(value2) ? string.Empty : value2.Trim();
            if (string.Equals(val1, val2, StringComparison.Ordinal))
            {
                return true;
            }
            string normVal1 = val1.Replace(',', '.');
            string normVal2 = val2.Replace(',', '.');
            bool isParsed1 = double.TryParse(normVal1, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double num1);
            bool isParsed2 = double.TryParse(normVal2, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double num2);
            if (isParsed1 && isParsed2)
            {
                return Math.Abs(num1 - num2) < 0.001;
            }
            return false;
        }
    } // Fin clase
} // Fin namespace