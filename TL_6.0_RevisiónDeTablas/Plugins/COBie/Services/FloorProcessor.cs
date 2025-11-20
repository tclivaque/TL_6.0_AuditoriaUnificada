using System;
using TL60_AuditoriaUnificada.Core;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using TL60_AuditoriaUnificada.Plugins.COBie.Models;
using TL60_AuditoriaUnificada.Models;
using System.Globalization;

namespace TL60_AuditoriaUnificada.Plugins.COBie.Services
{
    /// <summary>
    /// Procesador para Levels (FLOOR)
    /// </summary>
    public class FloorProcessor
    {
        private readonly Document _doc;
        private readonly List<ParameterDefinition> _definitions;
        private readonly GoogleSheetsService _sheetsService;
        private readonly CobieConfig _config;

        // Datos de Google Sheets "NIVELES"
        private Dictionary<string, FloorData> _floorData; // Key: "docTitle|levelId"

        public FloorProcessor(
    Document doc,
    List<ParameterDefinition> definitions,
    GoogleSheetsService sheetsService,
    CobieConfig config)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _definitions = definitions ?? throw new ArgumentNullException(nameof(definitions));
            _sheetsService = sheetsService ?? throw new ArgumentNullException(nameof(sheetsService));
            _config = config ?? throw new ArgumentNullException(nameof(config));

            // Cargar datos de Google Sheets "NIVELES"
            LoadFloorDataFromGoogleSheets();
        }

        /// <summary>
        /// Carga datos de la hoja "NIVELES" desde Google Sheets
        /// Estructura: Columna A=ID, B=Name, C=Category, D=Description, E=Height, F=Document Title
        /// </summary>
        private void LoadFloorDataFromGoogleSheets()
        {
            _floorData = new Dictionary<string, FloorData>(StringComparer.OrdinalIgnoreCase);

            try
            {
                const string SHEET_NAME = "NIVELES";
                var data = _sheetsService.ReadSheet(_config.SpreadsheetId, SHEET_NAME);

                if (data == null || data.Count <= 1)
                {
                    // Hoja vacía o sin datos - continuar sin datos
                    return;
                }

                // Procesar filas (saltando encabezado en fila 0)
                for (int i = 1; i < data.Count; i++)
                {
                    var row = data[i];
                    if (row.Count < 6) continue; // Necesitamos 6 columnas

                    string id = row[0]?.ToString().Trim();
                    string name = row[1]?.ToString().Trim();
                    string category = row[2]?.ToString().Trim();
                    string description = row[3]?.ToString().Trim();
                    string height = row[4]?.ToString().Trim();
                    string docTitle = row[5]?.ToString().Trim();

                    if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(docTitle))
                    {
                        continue; // Saltamos filas sin ID o sin Document Title
                    }

                    // Key = "DocumentTitle|LevelId"
                    string key = $"{docTitle}|{id}";

                    _floorData[key] = new FloorData
                    {
                        DocumentTitle = docTitle,
                        LevelId = id,
                        Name = name,
                        Category = category,
                        Description = description,
                        Height = height
                    };
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al cargar datos de Google Sheets 'NIVELES': {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Procesa todos los Levels del documento
        /// </summary>
        public List<ElementData> ProcessFloors()
        {
            var results = new List<ElementData>();

            try
            {
                // Filtrar solo parámetros FLOOR
                var parametrosFloor = _definitions
                    .Where(d => d.Grupo.Equals("FLOOR", StringComparison.OrdinalIgnoreCase)
                             || d.Grupo.Equals("3. FLOOR", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                // Obtener todos los Levels
                FilteredElementCollector collector = new FilteredElementCollector(_doc);
                IList<Element> levels = collector.OfClass(typeof(Level))
                                                .WhereElementIsNotElementType()
                                                .ToElements();

                // Procesar cada nivel
                foreach (var elem in levels)
                {
                    Level level = elem as Level;
                    if (level == null) continue;

                    var elementData = ProcessSingleFloor(level, parametrosFloor);
                    results.Add(elementData);
                }

                return results;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al procesar FLOORS: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Procesa un Level individual
        /// </summary>
        private ElementData ProcessSingleFloor(Level level, List<ParameterDefinition> parametrosFloor)
        {
            var elementData = new ElementData
            {
                Element = level,
                ElementId = level.Id,  // ← FIX: Asignar ElementId
                GrupoCOBie = "FLOOR",
                Nombre = level.Name,
                Categoria = "Levels"
            };

            // Procesar cada parámetro
            foreach (var def in parametrosFloor)
            {
                try
                {
                    ProcessFloorParameter(level, def, elementData);
                }
                catch (Exception ex)
                {
                    elementData.Mensajes.Add($"Error en parámetro {def.NombreParametro}: {ex.Message}");
                }
            }

            // Calcular si COBie está completo
            elementData.CobieCompleto = elementData.ParametrosVacios.Count == 0;

            return elementData;
        }

        /// <summary>
        /// Procesa un parámetro individual de FLOOR
        /// </summary>
        private void ProcessFloorParameter(
            Level level,
            ParameterDefinition def,
            ElementData elementData)
        {
            // Buscar parámetro en el Level
            Parameter param = level.LookupParameter(def.NombreParametro);
            if (param == null || param.IsReadOnly)
            {
                return; // Parámetro no existe o es de solo lectura
            }

            // Obtener valor actual
            string valorActual = GetParameterValue(param);

            // Determinar valor corregido
            string valorCorregido = CalculateFloorParameterValue(level, def);

            // ========== CORREGIDO: Manejar parámetros sin implementación ==========
            if (valorCorregido == null)
            {
                // Parámetro sin implementación → Marcar como VACÍO
                if (string.IsNullOrWhiteSpace(valorActual))
                {
                    elementData.ParametrosVacios.Add(def.NombreParametro);
                    elementData.ParametrosActualizar[def.NombreParametro] = "n/a";
                }
                else
                {
                    // Tiene valor pero no hay lógica → Considerar correcto
                    elementData.ParametrosCorrectos[def.NombreParametro] = valorActual;
                }
                return;
            }

            // Comparar valores
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
                // ========== NUEVO: Registrar parámetros CORRECTOS ==========
                elementData.ParametrosCorrectos[def.NombreParametro] = valorActual;
            }
        }

        /// <summary>
        /// Calcula el valor que debe tener un parámetro FLOOR
        /// </summary>
        private string CalculateFloorParameterValue(Level level, ParameterDefinition def)
        {
            // Si tiene valor fijo, usarlo
            if (def.TieneValorFijo)
            {
                return def.ValorFijo;
            }

            // Intentar obtener datos desde Google Sheets "NIVELES"
            // Doble filtro: Título del documento + ID del Level
            string docTitle = _doc.Title;
            string levelId = level.Id.IntegerValue.ToString();
            string key = $"{docTitle}|{levelId}";

            FloorData floorData = null;
            if (_floorData.ContainsKey(key))
            {
                floorData = _floorData[key];
            }

            // Valores dinámicos según el nombre del parámetro
            switch (def.NombreParametro)
            {
                case "COBie.Floor.Name":
                    // Desde Google Sheets
                    if (floorData != null && !string.IsNullOrWhiteSpace(floorData.Name))
                    {
                        return floorData.Name;
                    }
                    return string.Empty;

                case "COBie.Floor.Category":
                    // Desde Google Sheets
                    if (floorData != null && !string.IsNullOrWhiteSpace(floorData.Category))
                    {
                        return floorData.Category;
                    }
                    return string.Empty;

                case "COBie.Floor.Description":
                    // Desde Google Sheets
                    if (floorData != null && !string.IsNullOrWhiteSpace(floorData.Description))
                    {
                        return floorData.Description;
                    }
                    return string.Empty;

                case "COBie.Floor.Elevation":
                    // ========== CORREGIDO: Convertir de PIES a METROS (SIN REDONDEAR) ==========
                    double elevationPies = level.Elevation;
                    double elevationMetros = UnitUtils.ConvertFromInternalUnits(elevationPies, UnitTypeId.Meters);
                    // NO redondear aquí - dejar que la comparación use tolerancia
                    return elevationMetros.ToString("0.####");

                case "COBie.Floor.Height":
                    // Desde Google Sheets
                    if (floorData != null && !string.IsNullOrWhiteSpace(floorData.Height))
                    {
                        return floorData.Height;
                    }
                    return string.Empty;

                default:
                    // Sin implementación
                    return null;
            }
        }

        /// <summary>
        /// Obtiene el valor de un parámetro como string
        /// ========== CORREGIDO: Ahora convierte Double de PIES a METROS ==========
        /// </summary>
        private string GetParameterValue(Parameter param)
        {
            if (param == null || !param.HasValue)
                return string.Empty;

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

                        // SIEMPRE convertir - Revit almacena en pies internamente
                        if (unitTypeId != null &&
                            (unitTypeId.Equals(UnitTypeId.SquareFeet) ||
                             unitTypeId.Equals(UnitTypeId.SquareMeters)))
                        {
                            // Área: pies² → m²
                            double valorConvertido = UnitUtils.ConvertFromInternalUnits(valorInterno, UnitTypeId.SquareMeters);
                            return valorConvertido.ToString("0.####", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            // Longitud: pies → metros
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
            catch
            {
                return string.Empty;
            }
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
        /// Datos de un Level desde Google Sheets "NIVELES"
        /// </summary>
        private class FloorData
        {
            public string DocumentTitle { get; set; }
            public string LevelId { get; set; }
            public string Name { get; set; }
            public string Category { get; set; }
            public string Description { get; set; }
            public string Height { get; set; }
        }
    }
}