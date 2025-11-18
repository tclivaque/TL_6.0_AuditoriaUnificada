using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using TL60_RevisionDeTablas.Plugins.COBie.Models;
using TL60_RevisionDeTablas.Models;
using System.Globalization;

namespace TL60_RevisionDeTablas.Plugins.COBie.Services
{
    /// <summary>
    /// Procesador para parámetros FACILITY (Project Information)
    /// </summary>
    public class FacilityProcessor
    {
        private readonly Document _doc;
        private readonly List<ParameterDefinition> _definitions;

        public FacilityProcessor(Document doc, List<ParameterDefinition> definitions)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _definitions = definitions ?? throw new ArgumentNullException(nameof(definitions));
        }

        /// <summary>
        /// Procesa los parámetros FACILITY del Project Information
        /// </summary>
        public ElementData ProcessFacility()
        {
            try
            {
                // Obtener Project Information
                var collector = new FilteredElementCollector(_doc)
                    .OfCategory(BuiltInCategory.OST_ProjectInformation);

                var projectInfo = collector.FirstOrDefault();
                if (projectInfo == null)
                {
                    return null;
                }

                var elementData = new ElementData
                {
                    Element = projectInfo,
                    ElementId = projectInfo.Id,  // ← FIX: Asignar ElementId
                    GrupoCOBie = "FACILITY",
                    Nombre = "Project Information",
                    Categoria = "Project Information"
                };

                // Obtener definiciones de parámetros FACILITY
                var parametrosFacility = _definitions
                    .Where(d => d.Grupo.Equals("FACILITY", StringComparison.OrdinalIgnoreCase)
                             || d.Grupo.Equals("2. FACILITY", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var def in parametrosFacility)
                {
                    try
                    {
                        ProcessFacilityParameter(projectInfo, def, elementData);
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
            catch (Exception ex)
            {
                throw new Exception($"Error al procesar FACILITY: {ex.Message}", ex);
            }
        }

        private void ProcessFacilityParameter(
            Element projectInfo,
            ParameterDefinition def,
            ElementData elementData)
        {
            Parameter param = projectInfo.LookupParameter(def.NombreParametro);
            if (param == null || param.IsReadOnly)
            {
                return;
            }

            string valorActual = GetParameterValue(param);
            string valorCorregido = CalculateFacilityParameterValue(def);

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
                // Registrar parámetros CORRECTOS
                elementData.ParametrosCorrectos[def.NombreParametro] = valorActual;
            }
        }

        private string CalculateFacilityParameterValue(ParameterDefinition def)
        {
            // Si tiene valor fijo, usarlo
            if (def.TieneValorFijo)
            {
                return def.ValorFijo; // Incluye COBie.Facility.Phase (valor fijo desde Google Sheets)
            }

            // Valores dinámicos según el nombre del parámetro
            switch (def.NombreParametro)
            {
                case "COBie.Facility.Name":
                    // Usar nombre del sitio de Revit
                    Parameter siteNameParam = GetProjectInfoParameter("Site Name");
                    return siteNameParam?.AsString() ?? string.Empty;

                case "COBie.Facility.SiteName":
                    // Obtener título del documento (nombre del archivo sin extensión)
                    string titulo = _doc.Title;

                    if (string.IsNullOrEmpty(titulo))
                    {
                        return string.Empty;
                    }

                    // Verificar que el título tenga al menos 23 caracteres
                    if (titulo.Length < 23)
                    {
                        return string.Empty;
                    }

                    // Extraer caracteres en posiciones 19-21 (índices basados en 0)
                    string subString = titulo.Substring(19, 3); // Posiciones 19, 20, 21

                    if (subString == "000")
                    {
                        return "000";
                    }
                    else
                    {
                        // Extraer caracteres en posiciones 20-22 (índices basados en 0)
                        return titulo.Substring(20, 3); // Posiciones 20, 21, 22
                    }

                default:
                    return null;
            }
        }

        private Parameter GetProjectInfoParameter(string paramName)
        {
            var collector = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_ProjectInformation);

            var projectInfo = collector.FirstOrDefault();
            return projectInfo?.LookupParameter(paramName);
        }

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