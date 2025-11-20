// Plugins/Tablas/ScheduleUpdateWriter.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TL60_AuditoriaUnificada.Models;
using Autodesk.Revit.UI;

namespace TL60_AuditoriaUnificada.Plugins.Tablas
{
    public class ScheduleUpdateWriter
    {
        private readonly Document _doc;
        private const string EMPRESA_PARAM_NAME = "EMPRESA";
        private const string EMPRESA_PARAM_VALUE = "RNG";

        private void Log(string message)
        {
            System.Diagnostics.Debug.WriteLine(message);
        }

        public ScheduleUpdateWriter(Document doc)
        {
            _doc = doc;
        }

        // --- MÉTODO ORQUESTADOR ---
        public ProcessingResult UpdateAll(List<ElementData> todosLosElementos)
        {
            var finalResult = new ProcessingResult { Exitoso = true };
            var resultMessage = new StringBuilder();
            var allErrors = new List<string>();

            // 1. Separar las listas
            var unidadesGlobales = todosLosElementos
                .Where(ed => ed.Categoria == "UNIDADES GLOBALES")
                .ToList();
            var tablasData = todosLosElementos
                .Where(ed => ed.Categoria != "UNIDADES GLOBALES")
                .ToList();

            // 2. Corregir Unidades Globales (Primera Transacción)
            if (unidadesGlobales.Any(u => u.AuditResults.Any(a => a.IsCorrectable)))
            {
                ProcessingResult resultUnits = this.UpdateProjectUnits(unidadesGlobales);
                resultMessage.AppendLine(resultUnits.Mensaje);
                if (!resultUnits.Exitoso)
                {
                    finalResult.Exitoso = false;
                    allErrors.AddRange(resultUnits.Errores);
                }
            }

            // 3. Corregir Tablas (Segunda Transacción)
            if (tablasData.Any(t => t.AuditResults.Any(a => a.IsCorrectable)))
            {
                if (resultMessage.Length > 0) resultMessage.AppendLine();

                ProcessingResult resultTables = this.UpdateSchedules(tablasData);
                resultMessage.AppendLine(resultTables.Mensaje);
                if (!resultTables.Exitoso)
                {
                    finalResult.Exitoso = false;
                    allErrors.AddRange(resultTables.Errores);
                }
            }

            // 4. Combinar resultados
            finalResult.Mensaje = resultMessage.ToString();
            finalResult.Errores = allErrors;

            if (string.IsNullOrWhiteSpace(finalResult.Mensaje))
            {
                finalResult.Mensaje = "No se encontraron elementos que requieran corrección.";
                finalResult.Exitoso = true;
            }

            return finalResult;
        }


        // --- MÉTODO DE CORRECCIÓN DE TABLAS (COMPLETO) ---
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
            int parcialSaltados = 0;
            int empresaCorregidos = 0;

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
                        if (elementData.Categoria == "UNIDADES GLOBALES")
                            continue;

                        if (elementData.ElementId == null || elementData.ElementId == ElementId.InvalidElementId)
                            continue;

                        ViewSchedule view = _doc.GetElement(elementData.ElementId) as ViewSchedule;
                        if (view == null) continue;

                        ScheduleDefinition definition = view.Definition;
                        bool tablaModificada = false;

                        // ==========================================================
                        // ===== 1. CORRECCIÓN DE RECLASIFICACIÓN (LÓGICA MEJORADA)
                        // ==========================================================
                        // Busca CUALQUIER auditoría que sea corregible y 
                        // contenga RenamingJobData en su Tag.
                        var reclassAudit = elementData.AuditResults.FirstOrDefault(a => a.IsCorrectable && a.Tag is RenamingJobData);

                        if (reclassAudit != null)
                        {
                            // Se encontró una auditoría de reclasificación
                            // (puede ser "VIEW NAME", "ACER MANUAL", "COBie.Room", etc.)
                            RenamingJobData jobData = (RenamingJobData)reclassAudit.Tag;

                            if (RenameAndReclassify(view, jobData, result.Errores))
                            {
                                // Contar si el trabajo incluía un nuevo nombre
                                if (jobData.NuevoNombre != null && view.Name != jobData.NuevoNombre)
                                {
                                    nombresCorregidos++;
                                }
                                tablasReclasificadas++; // Contar esto como una reclasificación
                                tablaModificada = true;
                            }
                        }

                        // --- 2. Corregir FILTROS ---
                        var filterAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "FILTRO" && a.IsCorrectable);
                        if (filterAudit != null && filterAudit.Tag is List<ScheduleFilterInfo> filtrosCorrectos)
                        {
                            if (WriteFilters(definition, filtrosCorrectos, result.Errores, elementData.Nombre))
                            {
                                filtrosCorregidos++;
                                tablaModificada = true;
                            }
                        }

                        // --- 3. Corregir CONTENIDO (Itemize) ---
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

                        // --- 4. Corregir INCLUDE LINKS ---
                        var linksAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "LINKS" && a.IsCorrectable);
                        if (linksAudit != null)
                        {
                            if (!definition.IncludeLinkedFiles)
                            {
                                definition.IncludeLinkedFiles = true;
                                linksIncluidos++;
                                tablaModificada = true;
                            }
                        }

                        // --- 5. Corregir COLUMNAS ---
                        var columnAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "COLUMNAS" && a.IsCorrectable);
                        if (columnAudit != null && columnAudit.Tag != null)
                        {
                            if (columnAudit.Tag is Dictionary<ScheduleField, string> headingsToFix)
                            {
                                if (WriteHeadings(headingsToFix, result.Errores, elementData.Nombre))
                                {
                                    columnasRenombradas += headingsToFix.Count;
                                    tablaModificada = true;
                                }
                            }
                            else if (columnAudit.Tag is List<ScheduleField> fieldsToHide)
                            {
                                if (HideColumns(fieldsToHide, result.Errores, elementData.Nombre))
                                {
                                    columnasOcultadas += fieldsToHide.Count;
                                    tablaModificada = true;
                                }
                            }
                        }

                        // --- 6. Corregir FORMATO PARCIAL (Lógica simple) ---
                        var parcialAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "FORMATO PARCIAL" && a.IsCorrectable);
                        if (parcialAudit != null && parcialAudit.Tag is ScheduleFieldId)
                        {
                            try
                            {
                                ScheduleFieldId fieldId = (ScheduleFieldId)parcialAudit.Tag;
                                ScheduleField field = definition.GetField(fieldId);

                                if (field != null && field.IsValidObject)
                                {
                                    Log($"  [6] Activando 'Use project settings' en campo PARCIAL...");
                                    FormatOptions options = field.GetFormatOptions();

                                    options.UseDefault = true;
                                    field.SetFormatOptions(options);

                                    parcialCorregidos++;
                                    tablaModificada = true;
                                    Log("  [6] OK 'Use project settings' activado");
                                }
                            }
                            catch (Exception ex)
                            {
                                string errorMsg = $"No se pudo activar 'Use project settings' en '{elementData.Nombre}': {ex.Message}";
                                result.Errores.Add(errorMsg);
                                parcialSaltados++;
                                Log($"  [6] ERROR: {ex.Message}");
                            }
                        }

                        // --- 7. Corregir PARÁMETRO EMPRESA ---
                        var empresaAudit = elementData.AuditResults.FirstOrDefault(a => a.AuditType == "PARÁMETRO EMPRESA" && a.IsCorrectable);
                        if (empresaAudit != null)
                        {
                            try
                            {
                                Parameter param = view.LookupParameter(EMPRESA_PARAM_NAME);

                                if (param == null && empresaParamElem != null)
                                {
                                    result.Errores.Add($"Error en tabla '{elementData.Nombre}': El parámetro 'EMPRESA' no está vinculado. Ejecute el Add-in 'Verificar Parámetro Empresa'.");
                                }
                                else if (param != null && !param.IsReadOnly)
                                {
                                    param.Set(EMPRESA_PARAM_VALUE);
                                    empresaCorregidos++;
                                    tablaModificada = true;
                                }
                                else if (param == null && empresaParamElem == null)
                                {
                                    result.Errores.Add($"Error en tabla '{elementData.Nombre}': No se encontró el Parámetro Compartido 'EMPRESA' en el proyecto.");
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
                    } // Fin del foreach

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
                    result.Mensaje = $"Error fatal en la transacción de tablas: {ex.Message}";
                    result.Errores.Add(ex.Message);
                }
            } // Fin del using Transaction

            // --- 9. Lógica de Mensaje Final ---
            if (result.Exitoso)
            {
                result.Mensaje = $"Corrección de tablas completa.\n\n" +
                                 $"Tablas únicas modificadas: {tablasCorregidas}\n\n" +
                                 $"Detalles:\n" +
                                 $"- Formatos 'PARCIAL' corregidos (Use project settings): {parcialCorregidos}\n" +
                                 $"- Nombres de tabla corregidos: {nombresCorregidos}\n" +
                                 $"- Tablas reclasificadas: {tablasReclasificadas}\n" +
                                 $"- Parámetros 'EMPRESA' corregidos: {empresaCorregidos}\n" +
                                 $"- Filtros corregidos: {filtrosCorregidos}\n" +
                                 $"- Contenidos (Itemize) corregidos: {contenidosCorregidos}\n" +
                                 $"- 'Include Links' activados: {linksIncluidos}\n" +
                                 $"- Encabezados renombrados: {columnasRenombradas}\n" +
                                 $"- Columnas ocultadas: {columnasOcultadas}";
            }
            else
            {
                if (result.Errores.Count > 0)
                {
                    result.Mensaje = $"Se encontraron {result.Errores.Count} errores during la corrección de tablas (la transacción fue revertida):\n\n" +
                                     string.Join("\n", result.Errores.Take(5));
                }
                else if (string.IsNullOrEmpty(result.Mensaje))
                {
                    result.Mensaje = "La corrección de tablas se ejecutó pero no se detectaron errores. La transacción fue revertida.";
                }
            }

            if (parcialSaltados > 0)
                result.Mensaje += $"\n- Formatos 'PARCIAL' saltados por error: {parcialSaltados}";

            return result;
        }

        #region Corrección de Unidades Globales

        public ProcessingResult UpdateProjectUnits(List<ElementData> unitsData)
        {
            var result = new ProcessingResult { Exitoso = false };
            int unitsCorregidas = 0;

            Log("=== CORRECCIÓN DE UNIDADES GLOBALES ===\n");

            using (Transaction trans = new Transaction(_doc, "Corregir Unidades Globales"))
            {
                try
                {
                    trans.Start();
                    Units projectUnits = _doc.GetUnits();

                    foreach (var unitData in unitsData)
                    {
                        if (unitData.Categoria != "UNIDADES GLOBALES")
                            continue;
                        var auditItem = unitData.AuditResults.FirstOrDefault(a => a.IsCorrectable);
                        if (auditItem == null || auditItem.Tag == null)
                            continue;
                        ForgeTypeId specTypeId = auditItem.Tag as ForgeTypeId;
                        if (specTypeId == null)
                            continue;
                        Log($"\n--- Corrigiendo {unitData.Nombre} ---");

                        try
                        {
                            FormatOptions newFormat = null;
                            if (unitData.Nombre == "Volume")
                            {
                                newFormat = new FormatOptions(UnitTypeId.CubicMeters, new ForgeTypeId());
                                Log("  Configurando: Cubic meters, 2 decimales, sin símbolo");
                            }
                            else if (unitData.Nombre == "Area")
                            {
                                newFormat = new FormatOptions(UnitTypeId.SquareMeters, new ForgeTypeId());
                                Log("  Configurando: Square meters, 2 decimales, sin símbolo");
                            }
                            else if (unitData.Nombre == "Length")
                            {
                                newFormat = new FormatOptions(UnitTypeId.Meters, new ForgeTypeId());
                                Log("  Configurando: Meters, 2 decimales, sin símbolo");
                            }

                            if (newFormat != null)
                            {
                                newFormat.Accuracy = 0.01;
                                newFormat.UseDefault = false;
                                projectUnits.SetFormatOptions(specTypeId, newFormat);
                                unitsCorregidas++;
                                Log($"  OK {unitData.Nombre} configurado correctamente");
                            }
                        }
                        catch (Exception ex)
                        {
                            result.Errores.Add($"Error al configurar {unitData.Nombre}: {ex.Message}");
                            Log($"  ERROR: {ex.Message}");
                        }
                    }

                    _doc.SetUnits(projectUnits);
                    if (result.Errores.Count == 0)
                    {
                        result.Exitoso = true;
                    }

                    trans.Commit();
                    Log("\n=== UNIDADES GLOBALES ACTUALIZADAS ===");
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    result.Exitoso = false;
                    result.Mensaje = $"Error fatal al actualizar unidades: {ex.Message}";
                    result.Errores.Add(ex.Message);
                    Log($"\n=== ERROR FATAL: {ex.Message} ===");
                }
            }

            result.Mensaje = result.Exitoso
                ? $"✅ Unidades globales corregidas: {unitsCorregidas}"
                : $"❌ Error al corregir unidades globales";
            return result;
        }

        #endregion

        #region Métodos Helper

        private bool RenameAndReclassify(ViewSchedule view, RenamingJobData jobData, List<string> errores)
        {
            try
            {
                // Solo renombrar si el jobData lo especifica
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
            return null;
        }

        #endregion
    }
}