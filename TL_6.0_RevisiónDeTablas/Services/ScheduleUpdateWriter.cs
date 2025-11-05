// Services/ScheduleUpdateWriter.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Services
{
    /// <summary>
    /// Escribe TODAS las correcciones (Filtros, Contenido, Columnas, Renombrado)
    /// </summary>
    public class ScheduleUpdateWriter
    {
        private readonly Document _doc;

        public ScheduleUpdateWriter(Document doc)
        {
            _doc = doc;
        }

        public ProcessingResult UpdateSchedules(List<ElementData> elementosData)
        {
            var result = new ProcessingResult { Exitoso = false };
            int tablasCorregidas = 0;
            int filtrosCorregidos = 0;
            int contenidosCorregidos = 0;
            int columnasCorregidas = 0;
            int tablasRenombradas = 0; // (NUEVO)

            using (Transaction trans = new Transaction(_doc, "Corregir Auditoría de Tablas"))
            {
                try
                {
                    trans.Start();

                    foreach (var elementData in elementosData)
                    {
                        if (elementData.ElementId == null || elementData.ElementId == ElementId.InvalidElementId)
                            continue;

                        ViewSchedule view = _doc.GetElement(elementData.ElementId) as ViewSchedule;
                        if (view == null) continue;

                        ScheduleDefinition definition = view.Definition;
                        bool tablaModificada = false;

                        // --- 1. Corregir FILTROS (Solo si es corregible) ---
                        var filterAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "FILTRO" && a.IsCorrectable);
                        if (filterAudit != null)
                        {
                            var filtrosACorregir = filterAudit.Tag as List<ScheduleFilterInfo>;
                            if (filtrosACorregir != null && WriteFilters(definition, filtrosACorregir, result.Errores, elementData.Nombre))
                            {
                                filtrosCorregidos++;
                                tablaModificada = true;
                            }
                        }

                        // --- 2. Corregir CONTENIDO (Itemize) ---
                        var contentAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "CONTENIDO" && a.IsCorrectable);
                        if (contentAudit != null)
                        {
                            if (!definition.IsItemized)
                            {
                                definition.IsItemized = true;
                                contenidosCorregidos++;
                                tablaModificada = true;
                            }
                        }

                        // --- 3. Corregir COLUMNAS (Encabezados) ---
                        var columnAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "COLUMNAS" && a.IsCorrectable);
                        if (columnAudit != null)
                        {
                            var headingsToFix = columnAudit.Tag as Dictionary<ScheduleField, string>;
                            if (headingsToFix != null && WriteHeadings(headingsToFix, result.Errores, elementData.Nombre))
                            {
                                columnasCorregidas += headingsToFix.Count;
                                tablaModificada = true;
                            }
                        }

                        // ==========================================================
                        // ===== (NUEVO) 4. Ejecutar RENOMBRADO DE TABLA =====
                        // ==========================================================
                        var renameAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "CLASIFICACIÓN" && a.IsCorrectable);
                        if (renameAudit != null)
                        {
                            var jobData = renameAudit.Tag as RenamingJobData;
                            if (jobData != null && RenameAndReclassify(view, jobData, result.Errores))
                            {
                                tablasRenombradas++;
                                tablaModificada = true;
                            }
                        }
                        // ==========================================================

                        if (tablaModificada)
                        {
                            tablasCorregidas++;
                        }
                    }

                    trans.Commit();
                    result.Exitoso = true;
                    result.Mensaje = $"Corrección completa.\n" +
                                     $"Tablas auditadas modificadas: {tablasCorregidas - tablasRenombradas}\n" +
                                     $"Tablas de soporte reclasificadas: {tablasRenombradas}\n\n" +
                                     $"Detalles:\n" +
                                     $"Filtros corregidos: {filtrosCorregidos}\n" +
                                     $"Contenidos corregidos: {contenidosCorregidos}\n" +
                                     $"Encabezados corregidos: {columnasCorregidas}";
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    result.Exitoso = false;
                    result.Mensaje = $"Error al escribir correcciones: {ex.Message}";
                    result.Errores.Add(ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// (NUEVO) Renombra la tabla y actualiza sus parámetros de clasificación
        /// </summary>
        private bool RenameAndReclassify(ViewSchedule view, RenamingJobData jobData, List<string> errores)
        {
            try
            {
                // 1. Renombrar la vista
                view.Name = jobData.NuevoNombre;

                // 2. Establecer GRUPO DE VISTA
                Parameter paramGrupo = view.LookupParameter("GRUPO DE VISTA");
                if (paramGrupo != null && !paramGrupo.IsReadOnly)
                {
                    paramGrupo.Set(jobData.NuevoGrupoVista);
                }

                // 3. Establecer SUBGRUPO DE VISTA
                Parameter paramSubGrupo = view.LookupParameter("SUBGRUPO DE VISTA");
                if (paramSubGrupo != null && !paramSubGrupo.IsReadOnly)
                {
                    paramSubGrupo.Set(jobData.NuevoSubGrupoVista);
                }

                return true;
            }
            catch (Exception ex)
            {
                errores.Add($"Error al reclasificar tabla '{view.Name}': {ex.Message}");
                return false;
            }
        }

        private bool WriteHeadings(Dictionary<ScheduleField, string> headingsToFix, List<string> errores, string nombreTabla)
        {
            try
            {
                foreach (var kvp in headingsToFix)
                {
                    ScheduleField field = kvp.Key;
                    string correctedHeading = kvp.Value;

                    if (field != null && field.IsValidObject && field.ColumnHeading != correctedHeading)
                    {
                        field.ColumnHeading = correctedHeading;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                errores.Add($"Error al escribir encabezados en '{nombreTabla}': {ex.Message}");
                return false;
            }
        }

        private bool WriteFilters(ScheduleDefinition definition, List<ScheduleFilterInfo> filtrosCorrectos, List<string> errores, string nombreTabla)
        {
            try
            {
                definition.ClearFilters();

                foreach (var filtroInfo in filtrosCorrectos)
                {
                    ScheduleField field = FindField(definition, filtroInfo.FieldName);
                    if (field == null)
                    {
                        // Si el campo no existe (ej. falta "EMPRESA"), no podemos añadir el filtro.
                        // El procesador ya lo marcó como Error/Corregir, así que aquí solo lo ignoramos.
                        continue;
                    }

                    ScheduleFilter newFilter = CreateScheduleFilter(field.FieldId, filtroInfo, errores, nombreTabla);
                    if (newFilter != null)
                    {
                        definition.AddFilter(newFilter);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                errores.Add($"Error al escribir filtros en '{nombreTabla}': {ex.Message}");
                return false;
            }
        }

        private ScheduleFilter CreateScheduleFilter(ScheduleFieldId fieldId, ScheduleFilterInfo filtroInfo, List<string> errores, string nombreTabla)
        {
            if (filtroInfo.Value == null &&
                filtroInfo.FilterType != ScheduleFilterType.HasValue &&
                filtroInfo.FilterType != ScheduleFilterType.HasNoValue)
            {
                errores.Add($"Valor de filtro nulo no compatible para '{filtroInfo.FieldName}' en tabla '{nombreTabla}'.");
                return null;
            }

            switch (filtroInfo.Value)
            {
                case string s:
                    return new ScheduleFilter(fieldId, filtroInfo.FilterType, s);
                case double d:
                    return new ScheduleFilter(fieldId, filtroInfo.FilterType, d);
                case int i:
                    return new ScheduleFilter(fieldId, filtroInfo.FilterType, (double)i);
                case ElementId id:
                    return new ScheduleFilter(fieldId, filtroInfo.FilterType, id);
                case null:
                    return new ScheduleFilter(fieldId, filtroInfo.FilterType);
                default:
                    errores.Add($"Valor de filtro no compatible ({filtroInfo.Value.GetType()}) para '{filtroInfo.FieldName}' en tabla '{nombreTabla}'.");
                    return null;
            }
        }

        private ScheduleField FindField(ScheduleDefinition definition, string fieldName)
        {
            // (NUEVO) Búsqueda robusta que ignora prefijos
            for (int i = 0; i < definition.GetFieldCount(); i++)
            {
                var field = definition.GetField(i);
                var name = field.GetName();
                if (name.Equals(fieldName, StringComparison.OrdinalIgnoreCase) ||
                    name.EndsWith($": {fieldName}", StringComparison.OrdinalIgnoreCase))
                {
                    return field;
                }
            }

            // Fallback para campos sin prefijo (ej. "EMPRESA" si se añade nuevo)
            for (int i = 0; i < definition.GetFieldCount(); i++)
            {
                var field = definition.GetField(i);
                var name = field.GetName();
                if (name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                {
                    return field;
                }
            }

            return null;
        }
    }
}