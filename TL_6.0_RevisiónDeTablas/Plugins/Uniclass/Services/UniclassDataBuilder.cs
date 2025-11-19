// Plugins/Uniclass/Services/UniclassDataBuilder.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Plugins.Uniclass.Services
{
    /// <summary>
    /// Constructor de filas de diagnóstico para Uniclass
    /// </summary>
    public class UniclassDataBuilder
    {
        public List<DiagnosticRow> BuildDiagnosticRows(List<ElementData> elementosData)
        {
            var rows = new List<DiagnosticRow>();

            foreach (var elementData in elementosData)
            {
                if (elementData.Element == null) continue;

                string idMostrar = elementData.ElementId?.IntegerValue.ToString() ?? "N/A";
                bool esSeleccionable = (elementData.ElementId != null && elementData.ElementId != ElementId.InvalidElementId);

                // PARÁMETROS A CORREGIR
                foreach (var paramActualizar in elementData.ParametrosActualizar)
                {
                    string nombreParam = paramActualizar.Key;
                    string valorCorregido = paramActualizar.Value;
                    string valorActual = GetParameterValueFromElement(elementData.Element, nombreParam);

                    var row = new DiagnosticRow
                    {
                        ElementId = elementData.ElementId,
                        IdMostrar = idMostrar,
                        GrupoCOBie = "UNICLASS", // Usar campo GrupoCOBie para mantener compatibilidad
                        AssemblyCode = elementData.CodigoIdentificacion,
                        Descripcion = elementData.Nombre,
                        NombreParametro = nombreParam,
                        ValorActual = valorActual,
                        ValorCorregido = valorCorregido,
                        Estado = EstadoParametro.Corregir,
                        EsSeleccionable = esSeleccionable
                    };
                    rows.Add(row);
                }

                // PARÁMETROS CORRECTOS
                foreach (var paramCorrecto in elementData.ParametrosCorrectos)
                {
                    string nombreParam = paramCorrecto.Key;
                    string valorActual = GetParameterValueFromElement(elementData.Element, nombreParam);

                    var row = new DiagnosticRow
                    {
                        ElementId = elementData.ElementId,
                        IdMostrar = idMostrar,
                        GrupoCOBie = "UNICLASS",
                        AssemblyCode = elementData.CodigoIdentificacion,
                        Descripcion = elementData.Nombre,
                        NombreParametro = nombreParam,
                        ValorActual = valorActual,
                        ValorCorregido = valorActual,
                        Estado = EstadoParametro.Correcto,
                        EsSeleccionable = esSeleccionable
                    };
                    rows.Add(row);
                }

                // ERRORES
                foreach (var mensaje in elementData.Mensajes)
                {
                    var row = new DiagnosticRow
                    {
                        ElementId = elementData.ElementId,
                        IdMostrar = idMostrar,
                        GrupoCOBie = "UNICLASS",
                        AssemblyCode = elementData.CodigoIdentificacion,
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

            // Ordenar filas
            rows = rows.OrderBy(r => GetEstadoOrder(r.Estado))
                      .ThenBy(r => r.ElementId?.IntegerValue ?? 0)
                      .ToList();

            // Asignar números de fila
            for (int i = 0; i < rows.Count; i++)
            {
                rows[i].NumeroFila = i + 1;
            }

            // Aplicar colores alternados
            ApplyAlternatingColorsByID(rows);

            return rows;
        }

        private string GetParameterValueFromElement(Element element, string parameterName)
        {
            if (element == null) return string.Empty;

            Parameter param = element.LookupParameter(parameterName);
            if (param == null || !param.HasValue) return string.Empty;

            return param.AsString() ?? string.Empty;
        }

        private void ApplyAlternatingColorsByID(List<DiagnosticRow> rows)
        {
            string currentID = null;
            bool useAlternateColor = false;

            foreach (var row in rows)
            {
                if (row.IdMostrar != currentID)
                {
                    currentID = row.IdMostrar;
                    useAlternateColor = !useAlternateColor;
                }

                System.Windows.Media.Color baseColor = GetBaseColorForEstado(row.Estado);

                if (useAlternateColor)
                {
                    baseColor = DarkenColor(baseColor, 15);
                }

                row.BackgroundColor = new System.Windows.Media.SolidColorBrush(baseColor);
            }
        }

        private System.Windows.Media.Color GetBaseColorForEstado(EstadoParametro estado)
        {
            switch (estado)
            {
                case EstadoParametro.Correcto:
                    return System.Windows.Media.Color.FromRgb(212, 237, 218); // Verde
                case EstadoParametro.Vacio:
                    return System.Windows.Media.Color.FromRgb(255, 243, 205); // Amarillo
                case EstadoParametro.Corregir:
                    return System.Windows.Media.Color.FromRgb(209, 236, 241); // Azul
                case EstadoParametro.Error:
                    return System.Windows.Media.Color.FromRgb(248, 215, 218); // Rojo
                default:
                    return System.Windows.Media.Colors.White;
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
            switch (estado)
            {
                case EstadoParametro.Corregir: return 1;
                case EstadoParametro.Vacio: return 2;
                case EstadoParametro.Error: return 3;
                case EstadoParametro.Correcto: return 4;
                default: return 99;
            }
        }
    }
}
