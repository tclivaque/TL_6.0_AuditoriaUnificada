// Plugins/Tablas/DiagnosticDataBuilder.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using MediaColor = System.Windows.Media.Color;
using MediaBrush = System.Windows.Media.SolidColorBrush;
using TL60_RevisionDeTablas.Models;

// (¡ESTA ES LA LÍNEA CORREGIDA!)
namespace TL60_RevisionDeTablas.Plugins.Tablas
{
    public class DiagnosticDataBuilder
    {
        public List<DiagnosticRow> BuildDiagnosticRows(List<ElementData> elementosData)
        {
            var rows = new List<DiagnosticRow>();

            foreach (var elementData in elementosData)
            {
                // (¡AÑADIDO!) Si el ElementId es inválido (como nuestro informe de "Tablas Faltantes"), 
                // no debe ser seleccionable.
                bool esSeleccionable = (elementData.ElementId != null && elementData.ElementId != ElementId.InvalidElementId);
                string idMostrar = esSeleccionable ? elementData.ElementId.IntegerValue.ToString() : "N/A";

                // Procesar cada resultado de auditoría
                foreach (var auditResult in elementData.AuditResults)
                {
                    var row = new DiagnosticRow
                    {
                        ElementId = elementData.ElementId,
                        IdMostrar = idMostrar,
                        Grupo = elementData.Categoria,
                        CodigoIdentificacion = elementData.CodigoIdentificacion,
                        Descripcion = elementData.Nombre,
                        NombreParametro = auditResult.AuditType,
                        ValorActual = auditResult.ValorActual,
                        ValorCorregido = auditResult.ValorCorregido,
                        Estado = auditResult.Estado,
                        Mensaje = auditResult.Mensaje,
                        EsSeleccionable = esSeleccionable // <-- Usar la variable
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

            ApplyAlternatingColorsByID(rows);

            return rows;
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

                MediaColor baseColor = GetBaseColorForEstado(row.Estado);

                if (useAlternateColor)
                {
                    baseColor = DarkenColor(baseColor, 15);
                }

                row.BackgroundColor = new MediaBrush(baseColor);
            }
        }

        private MediaColor GetBaseColorForEstado(EstadoParametro estado)
        {
            switch (estado)
            {
                case EstadoParametro.Correcto:
                    return MediaColor.FromRgb(212, 237, 218); // Verde claro
                case EstadoParametro.Vacio:
                    return MediaColor.FromRgb(255, 243, 205); // Amarillo claro
                case EstadoParametro.Corregir:
                    return MediaColor.FromRgb(209, 236, 241); // Azul claro
                case EstadoParametro.Error:
                    return MediaColor.FromRgb(248, 215, 218); // Rojo claro
                default:
                    return MediaColor.FromRgb(255, 255, 255); // Blanco
            }
        }

        private MediaColor DarkenColor(MediaColor color, byte amount)
        {
            byte r = (byte)Math.Max(0, color.R - amount);
            byte g = (byte)Math.Max(0, color.G - amount);
            byte b = (byte)Math.Max(0, color.B - amount);
            return MediaColor.FromRgb(r, g, b);
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