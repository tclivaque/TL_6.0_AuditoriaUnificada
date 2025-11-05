// Services/ScheduleProcessor.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Services
{
    public class ScheduleProcessor
    {
        private readonly Document _doc;
        private readonly GoogleSheetsService _sheetsService;

        private readonly Dictionary<string, string> _aliasEncabezados = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "DIGO", "CODIGO" },
            { "DESCRIP", "DESCRIPCION" },
            { "ACTIV", "ACTIVO" },
            { "DULO", "MODULO" },
            { "NIVE", "NIVEL" },
            { "AMBIEN", "AMBIENTE" },
            { "EJE", "EJES" },
            { "PARC", "PARCIAL" },
            { "UNID", "UNIDAD" },
            { "ID", "ID" }
        };

        private readonly List<string> _expectedHeadings = new List<string>
        {
            "CODIGO",
            "DESCRIPCION",
            "ACTIVO",
            "MODULO",
            "NIVEL",
            "AMBIENTE",
            "PARCIAL",
            "UNIDAD",
            "ID"
        };
        private const int _expectedHeadingCount = 9;


        public ScheduleProcessor(Document doc, GoogleSheetsService sheetsService)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _sheetsService = sheetsService ?? throw new ArgumentNullException(nameof(sheetsService));
        }

        // =================================================================
        // ===== MÉTODO #1: Auditoría de "Tablas de Metrados" =====
        // =================================================================

        /// <summary>
        /// Ejecuta la auditoría completa (Filtros, Columnas, Contenido)
        /// </summary>
        public ElementData ProcessSingleElement(ViewSchedule view)
        {
            var elementData = new ElementData
            {
                ElementId = view.Id,
                Element = view,
                Nombre = view.Name,
                Categoria = view.Category?.Name ?? "Tabla de Planificación"
            };

            ScheduleDefinition definition = view.Definition;

            string[] parts = view.Name.Split(new[] { " - " }, StringSplitOptions.None);
            elementData.CodigoIdentificacion = parts.Length > 0 ?
                parts[0].Trim() : view.Name;

            // --- Ejecutar Auditorías ---
            var auditFilter = ProcessFilters(definition, elementData.CodigoIdentificacion);
            var auditColumns = ProcessColumns(definition);
            var auditContent = ProcessContent(view, definition);

            elementData.AuditResults.Add(auditFilter);
            elementData.AuditResults.Add(auditColumns);
            elementData.AuditResults.Add(auditContent);

            // --- Almacenar datos para corrección (si los hay) ---
            if (auditFilter.Estado == EstadoParametro.Corregir)
            {
                // Solo 'EMPRESA' es corregible
                auditFilter.Tag = auditFilter.Tag; // El tag ya contiene la lista
            }

            if (auditColumns.Estado == EstadoParametro.Corregir)
            {

            }

            if (auditContent.Estado == EstadoParametro.Corregir)
            {
                auditContent.Tag = true; // Marcar para corrección
            }

            elementData.DatosCompletos = elementData.AuditResults.All(r => r.Estado == EstadoParametro.Correcto);

            return elementData;
        }


        #region Auditoría 1: FILTRO (NUEVA LÓGICA DE ERROR)

        private AuditItem ProcessFilters(ScheduleDefinition definition, string assemblyCode)
        {
            var item = new AuditItem
            {
                AuditType = "FILTRO",
                IsCorrectable = false // Por defecto no es corregible, solo Empresa lo activa
            };

            var filtrosActuales = definition.GetFilters().ToList();
            var filtrosCorrectosInfo = new List<ScheduleFilterInfo>(); // Para "Valor Correcto"

            // 1. Identificar filtros "pie forzado" existentes
            int mainAssemblyCodeFilterIndex = FindMAINAssemblyCodeFilterIndex(filtrosActuales, definition);
            int empresaFilterIndex = FindEmpresaFilterIndex(filtrosActuales, definition);

            bool assemblyCodeCorrecto = false;
            bool empresaCorrecta = false;

            // 2. Auditar Assembly Code
            if (mainAssemblyCodeFilterIndex == -1) // No se encontró
            {
                item.Estado = EstadoParametro.Error;
                item.Mensaje = "Error: No se encontró filtro de Assembly Code (o Material: Assembly Code).";
            }
            else
            {
                // Se encontró, verificarlo
                ScheduleFilter filter = filtrosActuales[mainAssemblyCodeFilterIndex];
                string valorActual = GetFilterValueString(filter);

                if (valorActual.Equals(assemblyCode, StringComparison.OrdinalIgnoreCase))
                {
                    // Coinciden, es correcto
                    assemblyCodeCorrecto = true;
                    filtrosCorrectosInfo.Add(GetFilterInfo(filter, definition));
                }
                else
                {
                    // No coinciden
                    item.Estado = EstadoParametro.Error;
                    item.Mensaje = $"Error: El valor del filtro AC ('{valorActual}') no coincide con el del nombre ('{assemblyCode}').";
                    // Añadir el filtro incorrecto al "Valor Correcto" solo para visualización
                    filtrosCorrectosInfo.Add(GetFilterInfo(filter, definition));
                }
            }

            // 3. Auditar Empresa
            string valorEmpresaCorrecto = "RNG";
            if (empresaFilterIndex == -1) // No se encontró
            {
                item.Estado = EstadoParametro.Corregir; // Corregible
                item.Mensaje = (item.Mensaje ?? "") + "\nError: No se encontró filtro EMPRESA.";
                filtrosCorrectosInfo.Add(new ScheduleFilterInfo { FieldName = "EMPRESA", FilterType = ScheduleFilterType.Equal, Value = valorEmpresaCorrecto });
            }
            else
            {
                // Se encontró, verificarlo
                ScheduleFilter filter = filtrosActuales[empresaFilterIndex];
                var valorActual = GetFilterValueString(filter);

                if (filter.FilterType == ScheduleFilterType.Equal &&
                    valorActual.Equals(valorEmpresaCorrecto, StringComparison.OrdinalIgnoreCase))
                {
                    // Está perfecto
                    empresaCorrecta = true;
                    filtrosCorrectosInfo.Add(GetFilterInfo(filter, definition));
                }
                else
                {
                    // Está mal (ej. "EMPRESA Equal MAL")
                    item.Estado = EstadoParametro.Corregir; // Corregible
                    item.Mensaje = (item.Mensaje ?? "") + $"\nError: Filtro EMPRESA ('{valorActual}') no es 'RNG'.";
                    filtrosCorrectosInfo.Add(new ScheduleFilterInfo { FieldName = "EMPRESA", FilterType = ScheduleFilterType.Equal, Value = valorEmpresaCorrecto });
                }
            }

            // 4. Añadir todos los demás filtros en su orden original
            for (int i = 0; i < filtrosActuales.Count; i++)
            {
                if (i != mainAssemblyCodeFilterIndex && i != empresaFilterIndex)
                {
                    filtrosCorrectosInfo.Add(GetFilterInfo(filtrosActuales[i], definition));
                }
            }

            // 5. Finalizar
            item.ValorActual = GetFiltersAsString(definition, filtrosActuales);
            item.ValorCorrecto = string.Join("\n", filtrosCorrectosInfo.Select(f => f.AsString()));

            // Si el estado no es Error o Corregir, es Correcto
            if (item.Estado == 0 && assemblyCodeCorrecto && empresaCorrecta)
            {
                item.Estado = EstadoParametro.Correcto;
                item.Mensaje = "Filtros correctos.";
            }

            // Solo si el estado es "Corregir" (no "Error"), marcamos como corregible
            if (item.Estado == EstadoParametro.Corregir)
            {
                item.IsCorrectable = true;
                item.Tag = filtrosCorrectosInfo; // Guardar la lista para el Writer
            }

            return item;
        }

        // (Lógica de búsqueda de índices sin cambios)
        private int FindMAINAssemblyCodeFilterIndex(List<ScheduleFilter> filters, ScheduleDefinition def)
        {
            for (int i = 0; i < filters.Count; i++)
            {
                var filter = filters[i];
                var fieldName = def.GetField(filter.FieldId)?.GetName();
                if (IsAssemblyCodeField(fieldName))
                {
                    return i; // Devuelve el PRIMERO que encuentra
                }
            }
            return -1;
        }

        private int FindEmpresaFilterIndex(List<ScheduleFilter> filters, ScheduleDefinition def)
        {
            for (int i = 0; i < filters.Count; i++)
            {
                var fieldName = def.GetField(filters[i].FieldId)?.GetName();
                if (IsEmpresaField(fieldName))
                {
                    return i;
                }
            }
            return -1;
        }

        // (Lógica de comprobación de nombres sin cambios)
        private bool IsAssemblyCodeField(string fieldName)
        {
            if (string.IsNullOrEmpty(fieldName)) return false;
            string fn = fieldName.Trim();

            return fn.Equals("Assembly Code", StringComparison.OrdinalIgnoreCase) ||
                   fn.EndsWith(": Assembly Code", StringComparison.OrdinalIgnoreCase) ||
                   fn.Equals("MATERIAL_ASSEMBLY CODE", StringComparison.OrdinalIgnoreCase) ||
                   fn.EndsWith(": MATERIAL_ASSEMBLY CODE", StringComparison.OrdinalIgnoreCase);
        }

        private bool IsEmpresaField(string fieldName)
        {
            if (string.IsNullOrEmpty(fieldName)) return false;
            string fn = fieldName.Trim();

            return fn.Equals("EMPRESA", StringComparison.OrdinalIgnoreCase) ||
                   fn.EndsWith(": EMPRESA", StringComparison.OrdinalIgnoreCase);
        }

        #endregion

        #region Auditoría 2: COLUMNAS (Sin cambios)

        private AuditItem ProcessColumns(ScheduleDefinition definition)
        {
            var item = new AuditItem
            {
                AuditType = "COLUMNAS",
                IsCorrectable = true
            };

            var actualHeadings = new List<string>();
            var visibleFields = new List<ScheduleField>();

            for (int i = 0; i < definition.GetFieldCount(); i++)
            {
                var field = definition.GetField(i);
                if (!field.IsHidden)
                {
                    actualHeadings.Add(field.ColumnHeading.ToUpper().Trim());
                    visibleFields.Add(field);
                }
            }

            if (actualHeadings.Count != _expectedHeadingCount)
            {
                item.Estado = EstadoParametro.Error;
                item.Mensaje = $"Error: Se esperaban {_expectedHeadingCount} columnas visibles, pero se encontraron {actualHeadings.Count}.";
                item.ValorActual = $"Total: {actualHeadings.Count}\n" + string.Join("\n", actualHeadings);
                item.ValorCorrecto = $"Total: {_expectedHeadingCount}\n" + string.Join("\n", _expectedHeadings);
                item.IsCorrectable = false;
                return item;
            }

            var correctedHeadings = new List<string>();
            var headingsToFix = new Dictionary<ScheduleField, string>();
            bool isCorrect = true;

            for (int i = 0; i < _expectedHeadingCount; i++)
            {
                string actual = actualHeadings[i];
                string expectedBase = _expectedHeadings[i];
                string corrected = actual;

                if (i == 5) // Posición 6 (AMBIENTE o EJE)
                {
                    string expectedAlt = "EJES";
                    if (actual == expectedBase || actual == expectedAlt)
                    {
                        corrected = actual;
                    }
                    else if (_aliasEncabezados.TryGetValue(actual, out string aliasValue))
                    {
                        corrected = aliasValue;
                    }
                    else
                    {
                        isCorrect = false;
                        corrected = expectedBase;
                    }
                }
                else // Lógica normal
                {
                    if (actual == expectedBase)
                    {
                        corrected = actual;
                    }
                    else if (_aliasEncabezados.TryGetValue(actual, out string aliasValue))
                    {
                        corrected = aliasValue;
                    }
                }

                if (i == 5)
                {
                    if (corrected != "AMBIENTE" && corrected != "EJES")
                    {
                        isCorrect = false;
                        corrected = "AMBIENTE";
                    }
                }
                else if (corrected != expectedBase)
                {
                    isCorrect = false;
                    corrected = expectedBase;
                }

                correctedHeadings.Add(corrected);

                if (actual != corrected)
                {
                    isCorrect = false;
                    headingsToFix[visibleFields[i]] = corrected;
                }
            }

            item.ValorActual = string.Join("\n", actualHeadings);
            item.ValorCorrecto = string.Join("\n", correctedHeadings);

            if (isCorrect)
            {
                item.Estado = EstadoParametro.Correcto;
                item.Mensaje = "Encabezados correctos.";
            }
            else
            {
                item.Estado = EstadoParametro.Corregir;
                item.Mensaje = $"Error: {headingsToFix.Count} encabezados necesitan corrección.";
                item.Tag = headingsToFix;
            }

            return item;
        }

        #endregion

        #region Auditoría 3: CONTENIDO (Sin cambios)

        private AuditItem ProcessContent(ViewSchedule view, ScheduleDefinition definition)
        {
            var item = new AuditItem
            {
                AuditType = "CONTENIDO",
                IsCorrectable = true
            };

            bool isItemized = definition.IsItemized;

            string actualItemizedStr = isItemized ? "Sí" : "No";

            item.ValorActual = $"Detallar cada ejemplar: {actualItemizedStr}";
            item.ValorCorrecto = $"Detallar cada ejemplar: Sí";

            if (isItemized)
            {
                item.Estado = EstadoParametro.Correcto;
                item.Mensaje = "Correcto.";
            }
            else
            {
                item.Estado = EstadoParametro.Corregir;
                item.Mensaje = "Error: 'Detallar cada ejemplar' está desactivado.";
                item.Tag = true; // Marcar para corrección
            }

            return item;
        }

        #endregion

        // =================================================================
        // ===== MÉTODO #2: Corrección de "Tablas de Soporte" =====
        // =================================================================

        /// <summary>
        /// (NUEVO) Crea un trabajo de corrección para tablas que no son de metrados
        /// </summary>
        public ElementData CreateRenamingJob(ViewSchedule view)
        {
            var elementData = new ElementData
            {
                ElementId = view.Id,
                Element = view,
                Nombre = view.Name,
                Categoria = view.Category?.Name ?? "Tabla de Planificación",
                CodigoIdentificacion = "N/A",
                DatosCompletos = false // Siempre requiere corrección
            };

            string nombreActual = view.Name;
            string grupoActual = view.LookupParameter("GRUPO DE VISTA")?.AsString() ?? "(Vacío)";

            string nombreCorregido = nombreActual.Replace("C.", "SOPORTE.");
            string grupoCorregido = "00 TRABAJO EN PROCESO - WIP";
            string subGrupoCorregido = "SOPORTE DE METRADOS";

            var jobData = new RenamingJobData
            {
                NuevoNombre = nombreCorregido,
                NuevoGrupoVista = grupoCorregido,
                NuevoSubGrupoVista = subGrupoCorregido
            };

            var auditItem = new AuditItem
            {
                AuditType = "CLASIFICACIÓN", // Nuevo tipo de auditoría
                IsCorrectable = true,
                Estado = EstadoParametro.Corregir,
                Mensaje = "Tabla de soporte mal clasificada.",
                ValorActual = $"Nombre: {nombreActual}\nGrupo: {grupoActual}",
                ValorCorregido = $"Nombre: {nombreCorregido}\nGrupo: {grupoCorregido}\nSubGrupo: {subGrupoCorregido}",
                Tag = jobData // Adjuntar los datos del trabajo
            };

            elementData.AuditResults.Add(auditItem);
            return elementData;
        }

        #region Helpers de Filtros

        private ScheduleFilterInfo GetFilterInfo(ScheduleFilter filter, ScheduleDefinition def)
        {
            var field = def.GetField(filter.FieldId);
            return new ScheduleFilterInfo
            {
                FieldName = field.GetName(),
                FilterType = filter.FilterType,
                Value = GetFilterValueObject(filter)
            };
        }

        private string GetFiltersAsString(ScheduleDefinition definition, IList<ScheduleFilter> filters)
        {
            if (filters == null || filters.Count == 0)
                return "(Sin Filtros)";

            var filterStrings = new List<string>();
            foreach (var filter in filters)
            {
                if (filter == null) continue;
                ScheduleField field = definition.GetField(filter.FieldId);
                if (field == null) continue;

                string fieldName = field.GetName();
                string condition = filter.FilterType.ToString();
                string value = GetFilterValueString(filter);
                filterStrings.Add($"{fieldName} {condition} {value}");
            }
            return string.Join("\n", filterStrings);
        }

        private string GetFilterValueString(ScheduleFilter filter)
        {
            if (filter.FilterType == ScheduleFilterType.HasValue || filter.FilterType == ScheduleFilterType.HasNoValue)
                return string.Empty;

            if (filter.IsStringValue)
                return filter.GetStringValue();

            if (filter.IsDoubleValue)
                return filter.GetDoubleValue().ToString();

            if (filter.IsElementIdValue)
                return filter.GetElementIdValue().IntegerValue.ToString();

            return "(valor no legible)";
        }

        private object GetFilterValueObject(ScheduleFilter filter)
        {
            if (filter.FilterType == ScheduleFilterType.HasValue || filter.FilterType == ScheduleFilterType.HasNoValue)
                return null;

            if (filter.IsStringValue)
                return filter.GetStringValue();

            if (filter.IsDoubleValue)
                return filter.GetDoubleValue();

            if (filter.IsElementIdValue)
                return filter.GetElementIdValue();

            return null;
        }

        #endregion
    }
}