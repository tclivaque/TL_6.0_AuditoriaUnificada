// Plugins/Uniclass/Services/UniclassParameterWriter.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Plugins.Uniclass.Services
{
    /// <summary>
    /// Servicio para escribir parámetros Uniclass en tipos de elementos
    /// </summary>
    public class UniclassParameterWriter
    {
        private readonly Document _doc;

        public UniclassParameterWriter(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        /// <summary>
        /// Actualiza todos los parámetros Uniclass según los datos de auditoría
        /// </summary>
        public ProcessingResult UpdateParameters(List<ElementData> elementosData)
        {
            var result = new ProcessingResult { Exitoso = false };
            int tiposCorregidos = 0;
            int parametrosCorregidos = 0;

            using (Transaction trans = new Transaction(_doc, "Corregir Parámetros Uniclass"))
            {
                try
                {
                    trans.Start();

                    foreach (var elementData in elementosData)
                    {
                        if (elementData.Categoria != "UNICLASS")
                            continue;

                        if (elementData.ElementId == null || elementData.ElementId == ElementId.InvalidElementId)
                            continue;

                        ElementType tipo = _doc.GetElement(elementData.ElementId) as ElementType;
                        if (tipo == null) continue;

                        bool tipoModificado = false;

                        // Corregir parámetros que necesitan actualización
                        foreach (var paramActualizar in elementData.ParametrosActualizar)
                        {
                            string nombreParam = paramActualizar.Key;
                            string valorCorrecto = paramActualizar.Value;

                            Parameter param = tipo.LookupParameter(nombreParam);
                            if (param != null && !param.IsReadOnly)
                            {
                                try
                                {
                                    param.Set(valorCorrecto);
                                    parametrosCorregidos++;
                                    tipoModificado = true;
                                }
                                catch (Exception ex)
                                {
                                    result.Errores.Add($"Error al escribir '{nombreParam}' en tipo '{tipo.Name}': {ex.Message}");
                                }
                            }
                        }

                        if (tipoModificado)
                        {
                            tiposCorregidos++;
                        }
                    }

                    if (result.Errores.Count == 0)
                    {
                        result.Exitoso = true;
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    result.Exitoso = false;
                    result.Mensaje = $"Error fatal en la transacción: {ex.Message}";
                    result.Errores.Add(ex.Message);
                }
            }

            // Mensaje final
            if (result.Exitoso)
            {
                result.Mensaje = $"Corrección completa.\n\n" +
                                 $"Tipos corregidos: {tiposCorregidos}\n" +
                                 $"Parámetros actualizados: {parametrosCorregidos}";
            }
            else
            {
                if (result.Errores.Count > 0)
                {
                    result.Mensaje = $"Se encontraron {result.Errores.Count} errores durante la corrección:\n\n" +
                                     string.Join("\n", result.Errores.Take(5));
                }
            }

            return result;
        }
    }
}
