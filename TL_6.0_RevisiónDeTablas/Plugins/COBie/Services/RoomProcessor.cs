using System;
using TL60_AuditoriaUnificada.Core;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using TL60_AuditoriaUnificada.Plugins.COBie.Models;
using TL60_AuditoriaUnificada.Models;
using System.Globalization;

namespace TL60_AuditoriaUnificada.Plugins.COBie.Services
{
    /// <summary>
    /// Procesador para Rooms (SPACE)
    /// </summary>
    public class RoomProcessor
    {
        private readonly Document _doc;
        private readonly List<ParameterDefinition> _definitions;
        private readonly GoogleSheetsService _sheetsService;
        private readonly CobieConfig _config;

        // Datos de Google Sheets "AMBIENTES ÚNICOS"
        private Dictionary<string, string> _ambientesUnicos; // Col D (Category)
        private Dictionary<string, string> _ambientesSpaceName; // Col B (Name)
        private Dictionary<string, string> _ambientesDescription; // Col C (Description)
        // ===== CAMBIO AQUÍ: Nuevo diccionario para UsableHeight =====
        private Dictionary<string, string> _ambientesUsableHeight; // Col E (UsableHeight)

        public RoomProcessor(
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
        /// Inicializa cargando datos de Google Sheets
        /// </summary>
        public void Initialize()
        {
            LoadAmbientesUnicosFromGoogleSheets();
        }

        /// <summary>
        /// Carga datos de la hoja "AMBIENTES ÚNICOS" desde Google Sheets
        /// </summary>
        private void LoadAmbientesUnicosFromGoogleSheets()
        {
            _ambientesUnicos = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _ambientesSpaceName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _ambientesDescription = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            // ===== CAMBIO AQUÍ: Inicializar nuevo diccionario =====
            _ambientesUsableHeight = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                const string SHEET_NAME = "AMBIENTES ÚNICOS";
                var data = _sheetsService.ReadSheet(_config.SpreadsheetId, SHEET_NAME);

                if (data == null || data.Count <= 1)
                {
                    return; // Hoja vacía o sin datos
                }

                // Procesar filas (saltando encabezado en fila 0)
                for (int i = 1; i < data.Count; i++)
                {
                    var row = data[i];
                    if (row.Count < 2) continue; // Necesitamos al menos 2 columnas

                    string roomName = row[0]?.ToString().Trim().ToUpper(); // Col A
                    string spaceName = row.Count > 1 ? row[1]?.ToString().Trim() : null; // Col B
                    string description = row.Count > 2 ? row[2]?.ToString().Trim() : null; // Col C
                    string spaceCategory = row.Count > 3 ? row[3]?.ToString().Trim() : null; // Col D

                    // ===== CAMBIO AQUÍ: Leer Columna E y normalizar comas =====
                    string usableHeight = row.Count > 4 ? row[4]?.ToString().Trim().Replace(",", ".") : null; // Col E

                    if (string.IsNullOrWhiteSpace(roomName)) continue;

                    // Guardar los 4 mapeos
                    if (!string.IsNullOrEmpty(spaceName))
                    {
                        _ambientesSpaceName[roomName] = spaceName;
                    }

                    if (!string.IsNullOrEmpty(description))
                    {
                        _ambientesDescription[roomName] = description;
                    }

                    if (!string.IsNullOrEmpty(spaceCategory))
                    {
                        _ambientesUnicos[roomName] = spaceCategory;
                    }

                    // ===== CAMBIO AQUÍ: Guardar UsableHeight =====
                    if (!string.IsNullOrEmpty(usableHeight))
                    {
                        _ambientesUsableHeight[roomName] = usableHeight;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al cargar datos de Google Sheets 'AMBIENTES ÚNICOS': {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Procesa todos los Rooms del modelo
        /// </summary>
        public List<ElementData> ProcessRooms()
        {
            var results = new List<ElementData>();

            try
            {
                // Obtener todos los Rooms
                FilteredElementCollector collector = new FilteredElementCollector(_doc);
                IList<Element> rooms = collector.OfCategory(BuiltInCategory.OST_Rooms)
                                                .WhereElementIsNotElementType()
                                                .ToElements();

                // Obtener definiciones de parámetros SPACE
                var parametrosSpace = _definitions
                    .Where(d => d.Grupo.Equals("SPACE", StringComparison.OrdinalIgnoreCase)
                             || d.Grupo.Equals("4. SPACE", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                // Procesar cada Room
                foreach (Element elem in rooms)
                {
                    Room room = elem as Room;
                    if (room == null || room.Area == 0) continue; // Ignorar rooms sin colocar

                    try
                    {
                        var elementData = ProcessSingleRoom(room, parametrosSpace);
                        results.Add(elementData);
                    }
                    catch (Exception ex)
                    {
                        var errorData = new ElementData
                        {
                            Element = room,
                            ElementId = room.Id,  // ← FIX: Asignar ElementId
                            GrupoCOBie = "SPACE",
                            Nombre = room.get_Parameter(BuiltInParameter.ROOM_NAME)?.AsString() ?? "Sin nombre",
                            Categoria = "Rooms"
                        };
                        errorData.Mensajes.Add($"Error al procesar Room: {ex.Message}");
                        results.Add(errorData);
                    }
                }

                return results;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al procesar ROOMS: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Procesa un Room individual
        /// </summary>
        private ElementData ProcessSingleRoom(Room room, List<ParameterDefinition> parametrosSpace)
        {
            var elementData = new ElementData
            {
                ElementId = room.Id, // ← CORREGIDO: Asignar ElementId
                Element = room,
                GrupoCOBie = "SPACE", // ← CORRECTO: Grupo SPACE
                Nombre = room.Name,
                Categoria = "Rooms"
            };

            // Procesar cada parámetro SPACE
            foreach (var def in parametrosSpace)
            {
                try
                {
                    ProcessSpaceParameter(room, def, elementData);
                }
                catch (Exception ex)
                {
                    elementData.Mensajes.Add($"Error en parámetro {def.NombreParametro}: {ex.Message}");
                }
            }

            // ========== VALIDAR SI ROOM EXISTE EN GOOGLE SHEETS ==========
            Parameter checkRoomNameParam = room.get_Parameter(BuiltInParameter.ROOM_NAME);
            string checkRoomName = checkRoomNameParam?.AsString()?.Trim().ToUpper();

            bool roomFoundInSheets = false;

            if (!string.IsNullOrEmpty(checkRoomName))
            {
                // Verificar si existe en al menos uno de los diccionarios
                // ===== CAMBIO AQUÍ: Añadido chequeo a _ambientesUsableHeight =====
                roomFoundInSheets = _ambientesUnicos.ContainsKey(checkRoomName) ||
                                   _ambientesDescription.ContainsKey(checkRoomName) ||
                                   _ambientesSpaceName.ContainsKey(checkRoomName) ||
                                   _ambientesUsableHeight.ContainsKey(checkRoomName);
            }

            if (!roomFoundInSheets && !string.IsNullOrEmpty(checkRoomName))
            {
                // Room no encontrado en Google Sheets
                elementData.Mensajes.Add($"Nombre de Room '{checkRoomName}' no encontrado en AMBIENTES ÚNICOS");

                // Mover parámetros de "A Corregir" a "Vacíos" (para que se pinten amarillo)
                var parametrosToMove = elementData.ParametrosActualizar.Keys.ToList();
                foreach (var paramName in parametrosToMove)
                {
                    if (string.IsNullOrEmpty(elementData.ParametrosActualizar[paramName]))
                    {
                        elementData.ParametrosVacios.Add(paramName);
                    }
                }
            }

            // Calcular si COBie está completo
            elementData.CobieCompleto = elementData.ParametrosVacios.Count == 0 &&
                                       elementData.Mensajes.Count == 0;

            return elementData;
        }

        /// <summary>
        /// Procesa un parámetro individual de SPACE
        /// </summary>
        private void ProcessSpaceParameter(
            Room room,
            ParameterDefinition def,
            ElementData elementData)
        {
            // Buscar parámetro en el Room
            Parameter param = room.LookupParameter(def.NombreParametro);
            if (param == null || param.IsReadOnly)
            {
                return; // Parámetro no existe o es de solo lectura
            }

            // Obtener valor actual
            string valorActual = GetParameterValue(param);

            // Determinar valor corregido
            string valorCorregido = CalculateSpaceParameterValue(room, def);

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
        /// Calcula el valor que debe tener un parámetro SPACE
        /// </summary>
        private string CalculateSpaceParameterValue(Room room, ParameterDefinition def)
        {
            // Si tiene valor fijo, usarlo
            if (def.TieneValorFijo)
            {
                // Intentar formatear como número si es posible
                if (double.TryParse(def.ValorFijo, out double numValue))
                {
                    return numValue.ToString("0.00");
                }
                return def.ValorFijo;
            }

            // Valores dinámicos según el nombre del parámetro
            switch (def.NombreParametro)
            {
                case "COBie.Space.Name":
                    // Buscar en Google Sheets "AMBIENTES ÚNICOS" columna B
                    Parameter spaceNameRoomParam = room.get_Parameter(BuiltInParameter.ROOM_NAME);
                    string spaceNameRoom = spaceNameRoomParam?.AsString()?.Trim().ToUpper();

                    if (!string.IsNullOrEmpty(spaceNameRoom) && _ambientesSpaceName.ContainsKey(spaceNameRoom))
                    {
                        return _ambientesSpaceName[spaceNameRoom];
                    }
                    return string.Empty;

                case "COBie.Space.Category":
                    // Buscar en Google Sheets "AMBIENTES ÚNICOS" columna D
                    Parameter categoryRoomParam = room.get_Parameter(BuiltInParameter.ROOM_NAME);
                    string categoryRoom = categoryRoomParam?.AsString()?.Trim().ToUpper();

                    if (!string.IsNullOrEmpty(categoryRoom) && _ambientesUnicos.ContainsKey(categoryRoom))
                    {
                        return _ambientesUnicos[categoryRoom];
                    }
                    return string.Empty;

                case "COBie.Space.Description":
                    // Buscar en Google Sheets "AMBIENTES ÚNICOS" columna C
                    Parameter descRoomParam = room.get_Parameter(BuiltInParameter.ROOM_NAME);
                    string descRoom = descRoomParam?.AsString()?.Trim().ToUpper();

                    if (!string.IsNullOrEmpty(descRoom) && _ambientesDescription.ContainsKey(descRoom))
                    {
                        return _ambientesDescription[descRoom];
                    }
                    return string.Empty;

                case "COBie.Space.RoomTag":
                    // Nombre de habitación
                    Parameter roomNameParam = room.get_Parameter(BuiltInParameter.ROOM_NAME);
                    return roomNameParam?.AsString() ?? string.Empty;

                case "COBie.Space.GrossArea":
                    // SIEMPRE valor fijo "0" (no usar room.Area)
                    // Si hay valor fijo en Google Sheets, ya se usó arriba
                    return "0";

                case "COBie.Space.NetArea":
                    // ========== CORREGIDO: Convertir de PIES² a METROS² ==========
                    double areaPiesCuadrados = room.Area;
                    double areaMetrosCuadrados = UnitUtils.ConvertFromInternalUnits(areaPiesCuadrados, UnitTypeId.SquareMeters);
                    return areaMetrosCuadrados.ToString("0.####");

                // ===== CAMBIO AQUÍ: Nueva lógica para UsableHeight =====
                case "COBie.Space.UsableHeight":
                    // Buscar en Google Sheets "AMBIENTES ÚNICOS" columna E
                    Parameter heightRoomParam = room.get_Parameter(BuiltInParameter.ROOM_NAME);
                    string heightRoom = heightRoomParam?.AsString()?.Trim().ToUpper();

                    if (!string.IsNullOrEmpty(heightRoom) && _ambientesUsableHeight.ContainsKey(heightRoom))
                    {
                        return _ambientesUsableHeight[heightRoom];
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
    }
}