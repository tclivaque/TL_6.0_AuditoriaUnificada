using Autodesk.Revit.DB;
using TL60_RevisionDeTablas.Plugins.COBie.Models;
using TL60_RevisionDeTablas.Models;
using System;
using System.Collections.Generic;
using System.Globalization; // Necesario para CultureInfo
using System.Linq;

namespace TL60_RevisionDeTablas.Plugins.COBie.Services
{
    public class DiagnosticDataBuilder
    {
        public List<DiagnosticRow> BuildDiagnosticRows(List<ElementData> elementosData)
        {
            var rows = new List<DiagnosticRow>();

            foreach (var elementData in elementosData)
            {
                // Asegurarse de que elementData.Element no sea null para usarlo después
                if (elementData.Element == null) continue; // Saltar si no hay elemento asociado

                string idMostrar = elementData.ElementId?.IntegerValue.ToString() ?? "N/A";
                bool esSeleccionable = IsElementSelectable(elementData.GrupoCOBie, elementData.ElementId);

                // PARÁMETROS A CORREGIR (Azul)
                foreach (var paramActualizar in elementData.ParametrosActualizar)
                {
                    string nombreParam = paramActualizar.Key;
                    string valorCorregido = paramActualizar.Value;
                    // Pasar elementData.Element que sabemos que no es null
                    string valorActual = GetParameterActualValue(elementData.Element, nombreParam, elementData.GrupoCOBie);

                    var row = new DiagnosticRow
                    {
                        ElementId = elementData.ElementId,
                        IdMostrar = idMostrar,
                        GrupoCOBie = elementData.GrupoCOBie,
                        AssemblyCode = elementData.AssemblyCode,
                        Descripcion = elementData.Nombre,
                        NombreParametro = nombreParam,
                        ValorActual = valorActual,
                        ValorCorregido = valorCorregido,
                        Estado = EstadoParametro.Corregir,
                        EsSeleccionable = esSeleccionable
                    };
                    rows.Add(row);
                }

                // PARÁMETROS CORRECTOS (Verde)
                foreach (var paramCorrecto in elementData.ParametrosCorrectos)
                {
                    string nombreParam = paramCorrecto.Key;
                    string valorActualCorrecto = paramCorrecto.Value; // Valor guardado al procesar
                                                                      // Volver a leer de Revit para mostrar el valor más reciente
                    string valorActualLeido = GetParameterActualValue(elementData.Element, nombreParam, elementData.GrupoCOBie);

                    var row = new DiagnosticRow
                    {
                        ElementId = elementData.ElementId,
                        IdMostrar = idMostrar,
                        GrupoCOBie = elementData.GrupoCOBie,
                        AssemblyCode = elementData.AssemblyCode,
                        Descripcion = elementData.Nombre,
                        NombreParametro = nombreParam,
                        ValorActual = valorActualLeido,
                        ValorCorregido = valorActualCorrecto,
                        Estado = EstadoParametro.Correcto,
                        EsSeleccionable = esSeleccionable
                    };
                    rows.Add(row);
                }

                // ERRORES (Rojo) - Solo si hay mensajes reales
                foreach (var mensaje in elementData.Mensajes)
                {
                    var row = new DiagnosticRow
                    {
                        ElementId = elementData.ElementId,
                        IdMostrar = idMostrar,
                        GrupoCOBie = elementData.GrupoCOBie,
                        AssemblyCode = elementData.AssemblyCode,
                        Descripcion = elementData.Nombre,
                        NombreParametro = "ERROR",
                        ValorActual = string.Empty,
                        ValorCorregido = string.Empty,
                        Estado = EstadoParametro.Error,
                        Mensaje = mensaje,
                        EsSeleccionable = esSeleccionable
                    };
                    rows.Add(row);
                }
            }

            rows = rows.OrderBy(r => GetEstadoOrder(r.Estado))
                      .ThenBy(r => GetGrupoOrder(r.GrupoCOBie))
                      .ThenBy(r => r.ElementId?.IntegerValue ?? 0)
                      .ToList();

            for (int i = 0; i < rows.Count; i++) rows[i].NumeroFila = i + 1;

            ApplyAlternatingColorsByID(rows);
            return rows;
        }

        /// <summary>
        /// Obtiene el valor actual real de un parámetro desde el elemento.
        /// Busca primero en instancia, luego en tipo.
        /// </summary>
        private string GetParameterActualValue(Element element, string parameterName, string grupoCOBie)
        {
            // element ya está validado como no null por el llamador
            Parameter param = null;
            Document doc = element.Document;

            // 1. Intentar leer de la instancia
            param = element.LookupParameter(parameterName);

            // 2. Si no se encontró en instancia O si explícitamente es TYPE, buscar en el tipo
            if (param == null || grupoCOBie == "TYPE")
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    ElementType elementType = doc.GetElement(typeId) as ElementType;
                    if (elementType != null)
                    {
                        // Si era TYPE, sobreescribir; si era null, asignar.
                        param = elementType.LookupParameter(parameterName) ?? param;
                    }
                }
            }

            // Si no se encontró en ningún lado
            if (param == null) return string.Empty;

            // Convertir valor a string
            return GetParameterValueAsString(param);
        }

        /// <summary>
        /// Helper para convertir un valor de parámetro a string (incluye unidades)
        /// </summary>
        private string GetParameterValueAsString(Parameter param)
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
                        double val = param.AsDouble();
                        ForgeTypeId unit = param.GetUnitTypeId();
                        if (unit != null && (unit.Equals(UnitTypeId.SquareFeet) || unit.Equals(UnitTypeId.SquareMeters)))
                            // Usar CultureInfo.InvariantCulture para asegurar punto decimal
                            return UnitUtils.ConvertFromInternalUnits(val, UnitTypeId.SquareMeters).ToString("0.####", CultureInfo.InvariantCulture);
                        else
                            // Usar CultureInfo.InvariantCulture para asegurar punto decimal
                            return UnitUtils.ConvertFromInternalUnits(val, UnitTypeId.Meters).ToString("0.####", CultureInfo.InvariantCulture);
                    case StorageType.ElementId:
                        ElementId id = param.AsElementId();
                        if (id != null && id != ElementId.InvalidElementId)
                        {
                            // ===== CAMBIO AQUÍ: Usar param.Element.Document =====
                            Element refElement = param.Element?.Document?.GetElement(id);
                            return refElement?.Name ?? string.Empty;
                        }
                        return string.Empty;
                    default:
                        return param.AsValueString() ?? string.Empty; // Fallback
                }
            }
            catch { return string.Empty; } // Fallback en caso de error
        }


        private bool IsElementSelectable(string grupoCOBie, ElementId elementId)
        {
            if (elementId == null || elementId == ElementId.InvalidElementId) return false;
            string grupo = grupoCOBie?.ToUpper();
            return !(grupo == "FACILITY" || string.IsNullOrEmpty(grupo));
        }

        private void ApplyAlternatingColorsByID(List<DiagnosticRow> rows)
        {
            string currentID = null;
            bool useAlternateColor = false;
            foreach (var row in rows)
            {
                if (row.IdMostrar != currentID) { currentID = row.IdMostrar; useAlternateColor = !useAlternateColor; }
                System.Windows.Media.Color baseColor = GetBaseColorForEstado(row.Estado);
                if (useAlternateColor) baseColor = DarkenColor(baseColor, 15);
                row.BackgroundColor = new System.Windows.Media.SolidColorBrush(baseColor);
            }
        }

        private System.Windows.Media.Color GetBaseColorForEstado(EstadoParametro estado)
        {
            switch (estado)
            {
                case EstadoParametro.Correcto: return System.Windows.Media.Color.FromRgb(212, 237, 218); // Verde
                case EstadoParametro.Vacio: return System.Windows.Media.Color.FromRgb(255, 243, 205);    // Amarillo
                case EstadoParametro.Corregir: return System.Windows.Media.Color.FromRgb(209, 236, 241); // Azul
                case EstadoParametro.Error: return System.Windows.Media.Color.FromRgb(248, 215, 218);    // Rojo
                default: return System.Windows.Media.Colors.White;
            }
        }

        private System.Windows.Media.Color DarkenColor(System.Windows.Media.Color color, byte amount)
        {
            byte r = (byte)Math.Max(0, color.R - amount);
            byte g = (byte)Math.Max(0, color.G - amount);
            byte b = (byte)Math.Max(0, color.B - amount);
            return System.Windows.Media.Color.FromRgb(r, g, b);
        }

        private int GetEstadoOrder(EstadoParametro estado)
        {
            switch (estado) { case EstadoParametro.Corregir: return 1; case EstadoParametro.Vacio: return 2; case EstadoParametro.Error: return 3; case EstadoParametro.Correcto: return 4; default: return 99; }
        }

        private int GetGrupoOrder(string grupo)
        {
            switch (grupo?.ToUpper()) { case "FACILITY": return 1; case "FLOOR": return 2; case "SPACE": return 3; case "TYPE": return 4; case "COMPONENT": return 5; case "SIN_ASSEMBLY": return 6; case "ERROR": return 7; default: return 99; }
        }

        public string GenerateStats(List<DiagnosticRow> rows)
        {
            int totalCorregir = rows.Count(r => r.Estado == EstadoParametro.Corregir);
            int totalVacio = rows.Count(r => r.Estado == EstadoParametro.Vacio);
            int totalError = rows.Count(r => r.Estado == EstadoParametro.Error);
            int totalCorrecto = rows.Count(r => r.Estado == EstadoParametro.Correcto);
            return $"Total: {rows.Count} | Correctos: {totalCorrecto} | Vacíos: {totalVacio} | A corregir: {totalCorregir} | Errores: {totalError}";
        }
    } // Fin clase
      // ===== CÓDIGO INACCESIBLE/ERRÓNEO ELIMINADO ('g;') =====
} // Fin namespace