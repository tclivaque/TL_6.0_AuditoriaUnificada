// Plugins/Tablas/ScheduleUpdateWriter.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Plugins.Tablas
{
    public class ScheduleUpdateWriter
    {
        private readonly Document _doc;
        private const string EMPRESA_PARAM_NAME = "EMPRESA";

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
            int tablasReclasificadas = 0;
            int nombresCorregidos = 0;
            int linksIncluidos = 0;
            int columnasRenombradas = 0;
            int columnasOcultadas = 0;
            int parcialCorregidos = 0;
            int empresaCorregidos = 0; // <-- (¡NUEVO!)

            // (¡NUEVO!) Encontrar el Parámetro "EMPRESA" una sola vez
            ParameterElement empresaParamElem = new FilteredElementCollector(_doc)
                .OfClass(typeof(ParameterElement))
                .Cast<ParameterElement>()
                .FirstOrDefault(e => e.Name.Equals(EMPRESA_PARAM_NAME, StringComparison.OrdinalIgnoreCase));

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

                        // --- 1. Ejecutar RENOMBRADO DE TABLA (VIEW NAME) ---
                        // (Lógica sin cambios)
                        var viewNameAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "VIEW NAME" && a.IsCorrectable);
                        if (viewNameAudit != null)
                        {
                            string nuevoNombre = viewNameAudit.Tag as string;
                            if (!string.IsNullOrEmpty(nuevoNombre) && view.Name != nuevoNombre)
                            {
                                try
                                {
                                    view.Name = nuevoNombre;
                                    nombresCorregidos++;
                                    tablaModificada = true;
                                }
                                catch (Exception ex) { result.Errores.Add($"Error al renombrar tabla '{elementData.Nombre}': {ex.Message}"); }
                            }
                        }

                        // --- 2. Ejecutar RENOMBRADO DE CLASIFICACIÓN (WIP, MANUAL, SOPORTE, COPIA) ---
                        // (¡MODIFICADO! Ahora llama a la nueva versión de RenameAndReclassify)
                        var renameAudit = elementData.AuditResults.FirstOrDefault(a =>
                            (a.AuditType.StartsWith("CLASIFICACIÓN") ||
                             a.AuditType == "MANUAL" ||
                             a.AuditType == "COPIA" ||
                             a.AuditType.StartsWith("WIP")) // <-- Añadido por si acaso
                            && a.IsCorrectable);

                        if (renameAudit != null)
                        {
                            var jobData = renameAudit.Tag as RenamingJobData;
                            if (jobData != null && RenameAndReclassify(view, jobData, result.Errores))
                            {
                                tablasReclasificadas++;
                                tablaModificada = true;
                            }
                        }

                        // --- 3. Corregir FILTROS ---
                        // (Lógica sin cambios)
                        var filterAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "FILTRO" && a.IsCorrectable);
                        if (filterAudit != null) { /* ... */ }

                        // --- 4. Corregir CONTENIDO (Itemize) ---
                        // (Lógica sin cambios)
                        var contentAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "CONTENIDO" && a.IsCorrectable);
                        if (contentAudit != null) { /* ... */ }

                        // --- 5. Corregir INCLUDE LINKS ---
                        // (Lógica sin cambios)
                        var linksAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "LINKS" && a.IsCorrectable);
                        if (linksAudit != null) { /* ... */ }

                        // --- 6. Corregir COLUMNAS (Renombrar u Ocultar) ---
                        // (Lógica sin cambios)
                        var columnAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "COLUMNAS" && a.IsCorrectable);
                        if (columnAudit != null && columnAudit.Tag != null) { /* ... */ }

                        // --- 7. Corregir FORMATO PARCIAL ---
                        // (Lógica sin cambios, ya era correcta)
                        var parcialAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "FORMATO PARCIAL" && a.IsCorrectable);
                        if (parcialAudit != null && parcialAudit.Tag is ScheduleFieldId) { /* ... */ }


                        // ==========================================================
                        // ===== 8. (¡NUEVO!) Corregir PARÁMETRO EMPRESA
                        // ==========================================================
                        var empresaAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "PARÁMETRO EMPRESA" && a.IsCorrectable);
                        if (empresaAudit != null && empresaParamElem != null)
                        {
                            try
                            {
                                // Intentar obtener el parámetro de la tabla
                                Parameter param = view.LookupParameter(EMPRESA_PARAM_NAME);
                                if (param == null)
                                {
                                    // Si no existe, intentamos "forzar" la vinculación (esto es complejo y puede fallar
                                    // si el parámetro no está vinculado a OST_Schedules en el proyecto)
                                    // La lógica de TL_1_VerificarParametroEmpresa es la ideal, pero
                                    // por ahora, solo intentaremos establecerlo si existe.
                                    result.Errores.Add($"Error en tabla '{elementData.Nombre}': El parámetro 'EMPRESA' existe en el proyecto pero no está vinculado a esta tabla.");
                                }
                                else if (!param.IsReadOnly)
                                {
                                    param.Set(EMPRESA_PARAM_VALUE);
                                    empresaCorregidos++;
                                    tablaModificada = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                result.Errores.Add($"Error al asignar 'EMPRESA' en tabla '{elementData.Nombre}': {ex.Message}");
                            }
                        }

                        if (tablaModificada)
                        {
                            tablasCorregidas++;
                        }
                    }

                    trans.Commit();
                    result.Exitoso = true;

                    result.Mensaje = $"Corrección completa.\n\n" +
                                     $"Tablas únicas modificadas: {tablasCorregidas}\n\n" +
                                     $"Detalles:\n" +
                                     $"- Nombres de tabla corregidos: {nombresCorregidos}\n" +
                                     $"- Tablas reclasificadas: {tablasReclasificadas}\n" +
                                     $"- Parámetros 'EMPRESA' corregidos: {empresaCorregidos}\n" + // <-- Añadido
                                     $"- Filtros corregidos: {filtrosCorregidos}\n" +
                                     $"- Formatos 'PARCIAL' corregidos: {parcialCorregidos}\n" +
                                     $"- Contenidos (Itemize) corregidos: {contenidosCorregidos}\n" +
                                     $"- 'Include Links' activados: {linksIncluidos}\n" +
                                     $"- Encabezados renombrados: {columnasRenombradas}\n" +
                                     $"- Columnas ocultadas: {columnasOcultadas}";
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
        /// (¡MODIFICADO!) Ahora también asigna SUBPARTIDA
        /// </summary>
        private bool RenameAndReclassify(ViewSchedule view, RenamingJobData jobData, List<string> errores)
        {
            try
            {
                if (jobData.NuevoNombre != null && view.Name != jobData.NuevoNombre)
                {
                    view.Name = jobData.NuevoNombre;
                }

                Parameter paramGrupo = view.LookupParameter("GRUPO DE VISTA");
                if (paramGrupo != null && !paramGrupo.IsReadOnly && jobData.NuevoGrupoVista != null)
                {
                    paramGrupo.Set(jobData.NuevoGrupoVista);
                }

                Parameter paramSubGrupo = view.LookupParameter("SUBGRUPO DE VISTA");
                if (paramSubGrupo != null && !paramSubGrupo.IsReadOnly && jobData.NuevoSubGrupoVista != null)
                {
                    paramSubGrupo.Set(jobData.NuevoSubGrupoVista);
                }

                // (¡NUEVO!) Asignar el parámetro SUBPARTIDA
                Parameter paramSubPartida = view.LookupParameter("SUBGRUPO DE VISTA_SUBPARTIDA");
                if (paramSubPartida != null && !paramSubPartida.IsReadOnly && jobData.NuevoSubGrupoVistaSubpartida != null)
                {
                    paramSubPartida.Set(jobData.NuevoSubGrupoVistaSubpartida);
                }

                return true;
            }
            catch (Exception ex)
            {
                errores.Add($"Error al reclasificar tabla '{view.Name}': {ex.Message}");
                return false;
            }
        }

        // ... (WriteHeadings, HideColumns, WriteFilters, CreateScheduleFilter, FindField sin cambios) ...
        #region Métodos Helper Sin Cambios
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
        private bool HideColumns(List<ScheduleField> fieldsToHide, List<string> errores, string nombreTabla)
        {
            try
            {
                foreach (var field in fieldsToHide)
                {
                    if (field != null && field.IsValidObject && !field.IsHidden)
                    {
                        field.IsHidden = true;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                errores.Add($"Error al ocultar columnas en '{nombreTabla}': {ex.Message}");
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
        #endregion
    }
}