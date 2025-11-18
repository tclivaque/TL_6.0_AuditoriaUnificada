using System;
using TL60_RevisionDeTablas.Core;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TL60_RevisionDeTablas.Plugins.COBie.Models;
using TL60_RevisionDeTablas.Models;
using Autodesk.Revit.DB.Architecture;  // Para Room
using Autodesk.Revit.DB.Mechanical;    // Para MEPCurve
using Autodesk.Revit.DB.Plumbing;      // Para Pipe
using Autodesk.Revit.DB.Electrical;    // Para Conduit
using System.Globalization;

namespace TL60_RevisionDeTablas.Plugins.COBie.Services
{
    /// <summary>
    /// Procesador para elementos MEP (TYPE y COMPONENT)
    /// </summary>
    public class ElementProcessor
    {
        private readonly Document _doc;
        private readonly List<ParameterDefinition> _definitions;
        private readonly GoogleSheetsService _sheetsService;
        private readonly CobieConfig _config;

        // Datos de Google Sheets "ASSEMBLY CODE COBIE"
        private Dictionary<string, AssemblyCodeData> _assemblyData;
        private List<string> _cobieHeaders;


        public ElementProcessor(
            Document doc,
            List<ParameterDefinition> definitions,
            GoogleSheetsService sheetsService,
            CobieConfig config)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _definitions = definitions ?? throw new ArgumentNullException(nameof(definitions));
            _sheetsService = sheetsService ?? throw new ArgumentNullException(nameof(sheetsService));
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        /// <summary>
        /// Obtiene el ElementType de cualquier tipo de elemento (FamilyInstance, Wall, Floor, etc.)
        /// </summary>
        private ElementId GetElementTypeId(Element elemento)
        {
            if (elemento == null) return null;

            // Para FamilyInstance (puertas, ventanas, equipos, etc.)
            if (elemento is FamilyInstance fi)
                return fi.GetTypeId();

            // Para Wall (muros)
            else if (elemento is Wall wall)
                return wall.GetTypeId();

            // Para Floor (pisos/losas)
            else if (elemento is Floor floor)
                return floor.GetTypeId();

            // Para Ceiling (cielorrasos)
            else if (elemento is Ceiling ceiling)
                return ceiling.GetTypeId();

            // Para Roof (techos)
            else if (elemento is RoofBase roof)
                return roof.GetTypeId();

            // Para Stairs (escaleras)
            else if (elemento is Stairs stairs)
                return stairs.GetTypeId();

            // Para Railing (barandas)
            else if (elemento is Railing railing)
                return railing.GetTypeId();

            // Para MEP elements (tuberías, ductos, etc.)
            else if (elemento is MEPCurve mepCurve)
                return mepCurve.GetTypeId();

            // Fallback genérico (debería funcionar para la mayoría)
            else
            {
                try
                {
                    return elemento.GetTypeId();
                }
                catch
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Inicializa cargando datos de Google Sheets
        /// </summary>
        public void Initialize()
        {
            LoadAssemblyCodeDataFromGoogleSheets();
        }

        /// <summary>
        /// Procesa todos los elementos MEP del modelo según las categorías configuradas
        /// </summary>
        public List<ElementData> ProcessElements()
        {
            var results = new List<ElementData>();

            try
            {
                // Obtener elementos de todas las categorías configuradas
                var elementos = GetElementsByCategories(_config.Categorias);

                // Procesar cada elemento
                foreach (var elemento in elementos)
                {
                    try
                    {
                        var elementDataList = ProcessSingleElement(elemento);
                        results.AddRange(elementDataList);
                    }
                    catch (Exception ex)
                    {
                        // Continuar con el siguiente elemento si hay error
                        var errorData = new ElementData
                        {
                            Element = elemento,
                            ElementId = elemento.Id,
                            GrupoCOBie = "ERROR",
                            Nombre = elemento.Name
                        };
                        errorData.Mensajes.Add($"Error al procesar elemento: {ex.Message}");
                        results.Add(errorData);
                    }
                }

                return results;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al procesar elementos: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Procesa un elemento individual - AHORA retorna List<ElementData>
        /// </summary>
        private List<ElementData> ProcessSingleElement(Element elemento)
        {
            var results = new List<ElementData>();

            // Obtener Type del elemento
            ElementId typeId = GetElementTypeId(elemento);
            if (typeId == null || typeId == ElementId.InvalidElementId)
            {
                return results;
            }

            ElementType elementType = _doc.GetElement(typeId) as ElementType;
            if (elementType == null)
            {
                return results;
            }

            // Buscar Assembly Code en el Type
            Parameter assemblyParam = elementType.LookupParameter("Assembly Code");
            string assemblyCode = assemblyParam?.AsString()?.Trim() ?? string.Empty;

            // Manejar elementos SIN Assembly Code válido
            if (string.IsNullOrEmpty(assemblyCode) || !_assemblyData.ContainsKey(assemblyCode))
            {
                return ProcessElementWithoutAssemblyCode(elemento, elementType);
            }

            AssemblyCodeData assemblyData = _assemblyData[assemblyCode];

            // Crear ElementData PARA TYPE
            var elementDataType = new ElementData
            {
                Element = elemento,
                ElementId = elemento.Id,
                GrupoCOBie = "TYPE",
                Nombre = elementType?.Name ?? "Sin Tipo",
                Categoria = elemento.Category?.Name ?? "Sin Categoría",
                AssemblyCode = assemblyCode
            };

            ProcessTypeParameters(elementType, assemblyData, elementDataType);

            if (elementDataType.ParametrosActualizar.Count > 0 ||
                elementDataType.ParametrosVacios.Count > 0 ||
                elementDataType.ParametrosCorrectos.Count > 0)
            {
                results.Add(elementDataType);
            }

            // Crear ElementData PARA COMPONENT
            var elementDataComponent = new ElementData
            {
                Element = elemento,
                ElementId = elemento.Id,
                GrupoCOBie = "COMPONENT",
                Nombre = elemento.Name,
                Categoria = elemento.Category?.Name ?? "Sin Categoría",
                AssemblyCode = assemblyCode
            };

            ProcessComponentParameters(elemento, elementType, assemblyData, elementDataComponent);

            if (elementDataComponent.ParametrosActualizar.Count > 0 ||
                elementDataComponent.ParametrosVacios.Count > 0 ||
                elementDataComponent.ParametrosCorrectos.Count > 0)
            {
                results.Add(elementDataComponent);
            }

            return results;
        }

        /// <summary>
        /// Procesa elementos SIN Assembly Code válido
        /// </summary>
        private List<ElementData> ProcessElementWithoutAssemblyCode(Element elemento, ElementType elementType)
        {
            var results = new List<ElementData>();

            // PROCESAR TYPE
            if (elementType != null)
            {
                var elementDataType = new ElementData
                {
                    Element = elemento,
                    ElementId = elemento.Id,
                    GrupoCOBie = "TYPE",
                    Nombre = elementType.Name ?? "Sin Tipo",
                    Categoria = elemento.Category?.Name ?? "Sin Categoría",
                    AssemblyCode = "SIN ASSEMBLY CODE"
                };

                foreach (Parameter param in elementType.Parameters)
                {
                    if (param == null || param.IsReadOnly) continue;
                    string paramName = param.Definition.Name;

                    if (paramName.StartsWith("COBie", StringComparison.OrdinalIgnoreCase))
                    {
                        string valorActual = GetParameterValue(param);
                        string valorEsperado = GetDefaultValueForParameterType(param);

                        if (AreValuesEqual(valorActual, valorEsperado))
                        {
                            elementDataType.ParametrosCorrectos[paramName] = valorActual;
                        }
                        else
                        {
                            elementDataType.ParametrosActualizar[paramName] = valorEsperado;
                            if (string.IsNullOrWhiteSpace(valorActual))
                            {
                                elementDataType.ParametrosVacios.Add(paramName);
                            }
                        }
                    }
                }

                if (elementDataType.ParametrosActualizar.Count > 0 ||
                    elementDataType.ParametrosVacios.Count > 0 ||
                    elementDataType.ParametrosCorrectos.Count > 0)
                {
                    results.Add(elementDataType);
                }
            }

            // PROCESAR COMPONENT
            var elementDataComponent = new ElementData
            {
                Element = elemento,
                ElementId = elemento.Id,
                GrupoCOBie = "COMPONENT",
                Nombre = elemento.Name,
                Categoria = elemento.Category?.Name ?? "Sin Categoría",
                AssemblyCode = "SIN ASSEMBLY CODE"
            };

            foreach (Parameter param in elemento.Parameters)
            {
                if (param == null || param.IsReadOnly) continue;
                string paramName = param.Definition.Name;

                if (paramName.StartsWith("COBie", StringComparison.OrdinalIgnoreCase))
                {
                    string valorActual = GetParameterValue(param);
                    string valorEsperado = GetDefaultValueForParameterType(param);

                    if (AreValuesEqual(valorActual, valorEsperado))
                    {
                        elementDataComponent.ParametrosCorrectos[paramName] = valorActual;
                    }
                    else
                    {
                        elementDataComponent.ParametrosActualizar[paramName] = valorEsperado;
                        if (string.IsNullOrWhiteSpace(valorActual))
                        {
                            elementDataComponent.ParametrosVacios.Add(paramName);
                        }
                    }
                }
            }

            if (elementDataComponent.ParametrosActualizar.Count > 0 ||
                elementDataComponent.ParametrosVacios.Count > 0 ||
                elementDataComponent.ParametrosCorrectos.Count > 0)
            {
                results.Add(elementDataComponent);
            }

            return results;
        }

        /// <summary>
        /// Procesa parámetros COBie.Type.*
        /// </summary>
        private void ProcessTypeParameters(
            ElementType elementType,
            AssemblyCodeData assemblyData,
            ElementData elementData)
        {
            var parametrosType = _definitions
                .Where(d => d.Grupo.Equals("TYPE", StringComparison.OrdinalIgnoreCase)
                         || d.Grupo.Equals("5. TYPE", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var def in parametrosType)
            {
                try
                {
                    ProcessTypeParameter(elementType, def, assemblyData, elementData);
                }
                catch (Exception ex)
                {
                    elementData.Mensajes.Add($"Error en parámetro TYPE {def.NombreParametro}: {ex.Message}");
                }
            }
        }

        private void ProcessTypeParameter(
            ElementType elementType,
            ParameterDefinition def,
            AssemblyCodeData assemblyData,
            ElementData elementData)
        {
            Parameter param = elementType.LookupParameter(def.NombreParametro);
            if (param == null || param.IsReadOnly) return;

            string valorActual = GetParameterValue(param);
            string valorCorregido = CalculateTypeParameterValue(elementType, def, assemblyData);

            if (valorCorregido == null) return;

            if (!AreValuesEqual(valorActual, valorCorregido))
            {
                elementData.ParametrosActualizar[def.NombreParametro] = valorCorregido;
                if (string.IsNullOrWhiteSpace(valorActual))
                {
                    elementData.ParametrosVacios.Add(def.NombreParametro);
                }
            }
            else
            {
                elementData.ParametrosCorrectos[def.NombreParametro] = valorActual;
            }
        }

        private string CalculateTypeParameterValue(
            ElementType elementType,
            ParameterDefinition def,
            AssemblyCodeData assemblyData)
        {
            if (def.TieneValorFijo) return def.ValorFijo;

            switch (def.NombreParametro)
            {
                case "COBie.Type.Name":
                    Parameter nameParam = elementType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                    return nameParam?.AsString() ?? string.Empty;
                default:
                    if (assemblyData.ParametrosCobie.TryGetValue(def.NombreParametro, out string value))
                    {
                        return value;
                    }
                    return null;
            }
        }

        /// <summary>
        /// Procesa parámetros COBie.Component.*
        /// </summary>
        private void ProcessComponentParameters(
            Element elemento,
            ElementType elementType,
            AssemblyCodeData assemblyData,
            ElementData elementData)
        {
            var parametrosComponent = _definitions
                .Where(d => d.Grupo.Equals("COMPONENT", StringComparison.OrdinalIgnoreCase)
                         || d.Grupo.Equals("6. COMPONENT", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var def in parametrosComponent)
            {
                try
                {
                    ProcessComponentParameter(elemento, elementType, def, assemblyData, elementData);
                }
                catch (Exception ex)
                {
                    elementData.Mensajes.Add($"Error en parámetro COMPONENT {def.NombreParametro}: {ex.Message}");
                }
            }
        }

        private void ProcessComponentParameter(
            Element elemento,
            ElementType elementType,
            ParameterDefinition def,
            AssemblyCodeData assemblyData,
            ElementData elementData)
        {
            Parameter param = elemento.LookupParameter(def.NombreParametro);
            if (param == null || param.IsReadOnly) return;

            string valorActual = GetParameterValue(param);
            string valorCorregido = CalculateComponentParameterValue(elemento, elementType, def, assemblyData);

            if (valorCorregido == null) return;

            if (!AreValuesEqual(valorActual, valorCorregido))
            {
                elementData.ParametrosActualizar[def.NombreParametro] = valorCorregido;
                if (string.IsNullOrWhiteSpace(valorActual))
                {
                    elementData.ParametrosVacios.Add(def.NombreParametro);
                }
            }
            else
            {
                elementData.ParametrosCorrectos[def.NombreParametro] = valorActual;
            }
        }

        private string CalculateComponentParameterValue(
            Element elemento,
            ElementType elementType,
            ParameterDefinition def,
            AssemblyCodeData assemblyData)
        {
            if (def.TieneValorFijo) return def.ValorFijo;

            switch (def.NombreParametro)
            {
                case "COBie.Component.Name":
                    string elementId = elemento.Id.IntegerValue.ToString();
                    string categoryName = elemento.Category?.Name ?? "NoCategory";
                    string guid = elemento.UniqueId;
                    return $"{elementId}_{categoryName}_{guid}";
                case "COBie.Component.Space":
                    Parameter ambienteParam = elemento.LookupParameter("AMBIENTE");
                    return ambienteParam?.AsString() ?? string.Empty;
                case "COBie.Component.Description":
                    Parameter elementoParam = elementType.LookupParameter("ELEMENTO");
                    return elementoParam?.AsString() ?? string.Empty;
                case "COBie.Component.AssetIdentifier":
                    Parameter activoParam = elemento.LookupParameter("ACTIVO");
                    return activoParam?.AsString() ?? string.Empty;
                default:
                    if (assemblyData.ParametrosCobie.TryGetValue(def.NombreParametro, out string value))
                    {
                        return value;
                    }
                    return null;
            }
        }

        /// <summary>
        /// Obtiene elementos por categorías
        /// </summary>
        private List<Element> GetElementsByCategories(List<string> categorias)
        {
            var elementos = new List<Element>();
            foreach (var categoriaStr in categorias)
            {
                try
                {
                    if (Enum.TryParse(categoriaStr, out BuiltInCategory categoria))
                    {
                        var collector = new FilteredElementCollector(_doc)
                            .OfCategory(categoria)
                            .WhereElementIsNotElementType();
                        elementos.AddRange(collector.ToElements());
                    }
                }
                catch { continue; }
            }
            return elementos;
        }

        /// <summary>
        /// Carga datos de Assembly Code desde Google Sheets
        /// </summary>
        private void LoadAssemblyCodeDataFromGoogleSheets()
        {
            _assemblyData = new Dictionary<string, AssemblyCodeData>(StringComparer.OrdinalIgnoreCase);
            _cobieHeaders = new List<string>();

            try
            {
                var data = _sheetsService.ReadSheet(_config.SpreadsheetId, _config.NombreHojaAssemblyCode);
                if (data == null || data.Count <= 1)
                {
                    throw new Exception($"La hoja '{_config.NombreHojaAssemblyCode}' está vacía");
                }

                var headerRow = data[0];
                var (startCol, endCol) = GoogleSheetsService.ParseColumnRange(_config.RangoColumnasCOBie);

                for (int col = startCol; col <= endCol && col < headerRow.Count; col++)
                {
                    string header = GoogleSheetsService.GetCellValue(headerRow, col);
                    if (!string.IsNullOrWhiteSpace(header)) _cobieHeaders.Add(header);
                }

                for (int i = 1; i < data.Count; i++)
                {
                    var row = data[i];
                    if (row.Count < 2) continue;
                    string assemblyCode = GoogleSheetsService.GetCellValue(row, 1).ToUpper();
                    if (string.IsNullOrWhiteSpace(assemblyCode)) continue;

                    var assemblyCodeData = new AssemblyCodeData
                    {
                        AssemblyCode = assemblyCode,
                        ParametrosCobie = new Dictionary<string, string>()
                    };

                    for (int col = startCol; col <= endCol && col < row.Count; col++)
                    {
                        if (col - startCol < _cobieHeaders.Count)
                        {
                            string header = _cobieHeaders[col - startCol];
                            string valor = GoogleSheetsService.GetCellValue(row, col);
                            assemblyCodeData.ParametrosCobie[header] = valor;
                        }
                    }
                    _assemblyData[assemblyCode] = assemblyCodeData;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al cargar datos de Assembly Code: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Obtiene valor de parámetro como string (incluye conversión de unidades)
        /// </summary>
        private string GetParameterValue(Parameter param)
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
                        double valorInterno = param.AsDouble();
                        ForgeTypeId unitTypeId = param.GetUnitTypeId();
                        if (unitTypeId != null &&
                            (unitTypeId.Equals(UnitTypeId.SquareFeet) ||
                             unitTypeId.Equals(UnitTypeId.SquareMeters)))
                        {
                            double valorConvertido = UnitUtils.ConvertFromInternalUnits(valorInterno, UnitTypeId.SquareMeters);
                            return valorConvertido.ToString("0.####", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            double valorConvertido = UnitUtils.ConvertFromInternalUnits(valorInterno, UnitTypeId.Meters);
                            return valorConvertido.ToString("0.####", CultureInfo.InvariantCulture);
                        }
                    case StorageType.ElementId:
                        ElementId id = param.AsElementId();
                        if (id != null && id != ElementId.InvalidElementId)
                        {
                            Element elem = _doc.GetElement(id);
                            return elem?.Name ?? string.Empty;
                        }
                        return string.Empty;
                    default:
                        return param.AsValueString() ?? string.Empty;
                }
            }
            catch { return string.Empty; }
        }

        private bool AssemblyCodeEsValido(string assemblyCode)
        {
            if (string.IsNullOrWhiteSpace(assemblyCode)) return false;
            string code = assemblyCode.Trim().ToUpper();
            if (code.StartsWith("C.")) return true;
            string[] palabrasClave = { "TAKEOFF", "INCLUIDO EN APU", "METRADO MANUAL", "NO APLICA" };
            foreach (var palabra in palabrasClave)
            {
                if (code.Contains(palabra)) return true;
            }
            return false;
        }

        // --- CLASE INTERNA (SOLO UNA DEFINICIÓN) ---
        private class AssemblyCodeData
        {
            public string AssemblyCode { get; set; }
            public Dictionary<string, string> ParametrosCobie { get; set; }
        }

        /// <summary>
        /// Compara valores considerando conversiones numéricas (0 == 0.00)
        /// CORREGIDO: Usa CultureInfo.InvariantCulture para ser insensible a coma/punto.
        /// </summary>
        private bool AreValuesEqual(string value1, string value2)
        {
            // Normalizar valores nulos o vacíos a string vacío
            string val1 = string.IsNullOrWhiteSpace(value1) ? string.Empty : value1.Trim();
            string val2 = string.IsNullOrWhiteSpace(value2) ? string.Empty : value2.Trim();

            // Comparación exacta de cadenas (ignorando mayúsculas/minúsculas)
            if (string.Equals(val1, val2, StringComparison.Ordinal))
            {
                return true;
            }

            // ===== CAMBIO AQUÍ: Normalizar separadores y usar CultureInfo.InvariantCulture =====
            // Reemplazar comas por puntos antes de intentar convertir
            string normVal1 = val1.Replace(',', '.');
            string normVal2 = val2.Replace(',', '.');

            // Intentar parsear ambos como números usando cultura invariante (espera '.')
            bool isParsed1 = double.TryParse(normVal1, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double num1);
            bool isParsed2 = double.TryParse(normVal2, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double num2);

            // Si ambos son números válidos, comparar con tolerancia
            if (isParsed1 && isParsed2)
            {
                // Tolerancia de 0.001 (ej: 1mm si las unidades son metros)
                return Math.Abs(num1 - num2) < 0.001;
            }

            // Si no son ambos números (o la comparación exacta falló), son diferentes
            return false;
        }

        /// <summary>
        /// Obtiene valor por defecto (ahora lee desde _config)
        /// </summary>
        private string GetDefaultValueForParameterType(Parameter param)
        {
            if (param == null) return _config.ValorDefectoString;
            switch (param.StorageType)
            {
                case StorageType.Integer:
                    return (param.Definition.ParameterType == ParameterType.YesNo)
                        ? _config.ValorDefectoIntegerYesNo
                        : _config.ValorDefectoInteger;
                case StorageType.Double:
                    return _config.ValorDefectoDouble;
                case StorageType.String:
                default:
                    return _config.ValorDefectoString;
            }
        }
    }
}