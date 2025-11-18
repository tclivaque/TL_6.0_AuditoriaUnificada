using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using TL60_RevisionDeTablas.Plugins.COBie.Models;
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Plugins.COBie.Services
{
    /// <summary>
    /// Escribe parámetros en elementos de Revit
    /// </summary>
    public class ParameterWriter
    {
        private readonly Document _doc;
        // Ya no necesita config ni definitions
        // private CobieConfig _config;
        // private List<ParameterDefinition> _definitions;

        // Constructor original
        public ParameterWriter(Document doc)
        {
            _doc = doc;
        }

        /// <summary>
        /// Escribe los parámetros corregidos en los elementos
        /// (Solo escribe lo que viene en elementosData.ParametrosActualizar)
        /// </summary>
        public ProcessingResult WriteParameters(List<ElementData> elementosData)
        {
            var result = new ProcessingResult { Exitoso = false, Mensaje = string.Empty };
            int parametrosEscritos = 0;
            int elementosProcesados = 0;

            try
            {
                using (Transaction trans = new Transaction(_doc, "Corregir Parámetros COBie"))
                {
                    trans.Start();
                    foreach (var elementData in elementosData)
                    {
                        if (elementData.ElementId == null || elementData.ElementId == ElementId.InvalidElementId) continue;
                        Element elemento = _doc.GetElement(elementData.ElementId);
                        if (elemento == null) continue;

                        ElementType elementType = null;
                        ElementId typeId = GetElementTypeId(elemento);
                        if (typeId != null && typeId != ElementId.InvalidElementId)
                        {
                            elementType = _doc.GetElement(typeId) as ElementType;
                        }

                        foreach (var paramActualizar in elementData.ParametrosActualizar)
                        {
                            string nombreParam = paramActualizar.Key;
                            string valorCorregido = paramActualizar.Value;
                            try
                            {
                                Parameter param = null;
                                // Buscar parámetro (en tipo o instancia según GrupoCOBie)
                                if (elementData.GrupoCOBie == "TYPE")
                                {
                                    if (elementType != null) param = elementType.LookupParameter(nombreParam);
                                }
                                else
                                {
                                    param = elemento.LookupParameter(nombreParam);
                                    if (param == null && elementType != null) param = elementType.LookupParameter(nombreParam);
                                }

                                if (param != null && !param.IsReadOnly)
                                {
                                    if (SetParameterValue(param, valorCorregido)) parametrosEscritos++;
                                    else result.Errores.Add($"No se pudo escribir {nombreParam} en {elementData.ElementId}");
                                }
                                else
                                {
                                    // No añadir error si el parámetro no existe, simplemente no se escribe
                                    // result.Errores.Add($"Parámetro {nombreParam} no encontrado/lectura en {elementData.ElementId}");
                                }
                            }
                            catch (Exception ex)
                            {
                                result.Errores.Add($"Error al escribir {nombreParam} ({elementData.ElementId}): {ex.Message}");
                            }
                        }
                        elementosProcesados++;
                    }

                    // Limpiar ParametrosVacios si se corrigió algo
                    foreach (var elementData in elementosData.Where(ed => ed.ParametrosActualizar.Any()))
                    {
                        // Si se escribió algo, ya no está vacío, pero mantenemos ParametrosCorrectos si existían
                        elementData.ParametrosVacios.Clear();
                        // Actualizar estado CobieCompleto si aplica
                        elementData.CobieCompleto = elementData.ParametrosVacios.Count == 0 && elementData.Mensajes.Count == 0;
                    }

                    MarkCobieCheckboxes(elementosData); // Marcar checkboxes si ahora está completo
                    trans.Commit();
                }
                result.Exitoso = true;
                result.Mensaje = $"✅ Se actualizaron {parametrosEscritos} parámetros en {elementosProcesados} elementos.";
            }
            catch (Exception ex)
            {
                result.Exitoso = false;
                result.Mensaje = $"❌ Error al escribir parámetros: {ex.Message}";
                result.Errores.Add(ex.Message);
            }
            return result;
        }

        // --- LOS MÉTODOS FillDefaultsForGroup, FillElementDefaults, FillRemainingCobieParams y GetDefaultValueForParameterType HAN SIDO ELIMINADOS ---

        /// <summary>
        /// Marca los checkboxes COBie
        /// </summary>
        private void MarkCobieCheckboxes(List<ElementData> elementosData)
        {
            foreach (var elementData in elementosData.Where(ed => ed.CobieCompleto)) // Solo marcar si está completo
            {
                // El resto de la lógica es igual...
                string checkboxParamName = null;
                Element targetElement = null;
                switch (elementData.GrupoCOBie?.ToUpper())
                {
                    case "FACILITY":
                    case "FLOOR":
                    case "SPACE":
                    case "COMPONENT":
                        checkboxParamName = "COBie";
                        targetElement = elementData.Element;
                        break;
                    case "TYPE":
                        checkboxParamName = "Type.COBie";
                        ElementId typeId = GetElementTypeId(elementData.Element);
                        if (typeId != null && typeId != ElementId.InvalidElementId) targetElement = _doc.GetElement(typeId);
                        break;
                    default: continue;
                }

                if (targetElement == null || string.IsNullOrEmpty(checkboxParamName)) continue;
                Parameter checkboxParam = targetElement.LookupParameter(checkboxParamName);
                if (checkboxParam == null || checkboxParam.IsReadOnly || checkboxParam.StorageType != StorageType.Integer) continue;

                try { checkboxParam.Set(1); } // Marcar como True
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"No se pudo marcar {checkboxParamName}: {ex.Message}"); }
            }
        }

        /// <summary>
        /// Establece el valor de un parámetro (incluye conversión de unidades)
        /// </summary>
        private bool SetParameterValue(Parameter param, string value)
        {
            try
            {
                if (value == null) value = string.Empty;
                // Si el valor es "n/a", para tipos numéricos lo tratamos como 0
                string numericValue = value.Equals("n/a", StringComparison.OrdinalIgnoreCase) ? "0" : value;

                switch (param.StorageType)
                {
                    case StorageType.String:
                        param.Set(value); // Para string, se escribe "n/a" tal cual si ese es el valor
                        return true;
                    case StorageType.Integer:
                        // Para YesNo, si el valor es "0" o "1", TryParse funciona. Si es "n/a", numericValue es "0".
                        if (int.TryParse(numericValue, out int intValue)) { param.Set(intValue); return true; }
                        return false;
                    case StorageType.Double:
                        if (double.TryParse(numericValue, out double doubleValue))
                        {
                            ForgeTypeId unitTypeId = param.GetUnitTypeId();
                            double valueInFeet;
                            if (unitTypeId != null && (unitTypeId.Equals(UnitTypeId.SquareFeet) || unitTypeId.Equals(UnitTypeId.SquareMeters)))
                                valueInFeet = UnitUtils.ConvertToInternalUnits(doubleValue, UnitTypeId.SquareMeters); // m² -> ft²
                            else
                                valueInFeet = UnitUtils.ConvertToInternalUnits(doubleValue, UnitTypeId.Meters); // m -> ft
                            param.Set(valueInFeet);
                            return true;
                        }
                        return false;
                    case StorageType.ElementId:
                        if (int.TryParse(numericValue, out int idValue) && idValue > 0) // Asegurarse que el ID sea válido
                        {
                            param.Set(new ElementId(idValue));
                            return true;
                        }
                        else if (numericValue == "0")
                        {
                            // Intentar poner ElementId.InvalidElementId si el valor es 0
                            param.Set(ElementId.InvalidElementId);
                            return true;
                        }
                        return false;
                    default:
                        return false;
                }
            }
            catch { return false; }
        }

        /// <summary>
        /// Obtiene el ElementTypeId de cualquier tipo de elemento
        /// </summary>
        private ElementId GetElementTypeId(Element elemento)
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
    } // Fin clase
} // Fin namespace