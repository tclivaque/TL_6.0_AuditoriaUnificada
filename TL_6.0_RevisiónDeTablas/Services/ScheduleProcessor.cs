// Services/ScheduleProcessor.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions; // (NUEVO)
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Services
{
    public class ScheduleProcessor
    {
        private readonly Document _doc;
        private readonly GoogleSheetsService _sheetsService;

        // (NUEVO) Regex para extraer Assembly Code
        private static readonly Regex _acRegex = new Regex(@"^C\.(\d{2}\.)+\d{2}");

        // (NUEVO) Lista de nombres WIP (de proyecto 5.0)
        private static readonly List<string> NOMBRES_WIP = new List<string>
        {
            "TL", "TITO", "PDONTADENEA", "ANDREA", "EFRAIN",
            "PROYECTOSBIM", "ASISTENTEBIM", "LUIS", "DIEGO", "JORGE", "MIGUEL"
        };

        private readonly Dictionary<string, string> _aliasEncabezados = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "DIGO", "CODIGO" }, { "DESCRIP", "DESCRIPCION" }, { "ACTIV", "ACTIVO" },
            { "DULO", "MODULO" }, { "NIVE", "NIVEL" }, { "AMBIEN", "AMBIENTE" },
            { "EJE", "EJES" }, { "PARC", "PARCIAL" }, { "UNID", "UNIDAD" }, { "ID", "ID" }
        };

        private readonly List<string> _expectedHeadings = new List<string>
        {
            "CODIGO", "DESCRIPCION", "ACTIVO", "MODULO", "NIVEL",
            "AMBIENTE", "PARCIAL", "UNIDAD", "ID"
        };
        private const int _expectedHeadingCount = 9;


        public ScheduleProcessor(Document doc, GoogleSheetsService sheetsService, ManualScheduleService manualScheduleService)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _sheetsService = sheetsService ?? throw new ArgumentNullException(nameof(sheetsService));
            // (MODIFICADO) El servicio G-Sheets se mantiene, el ManualScheduleService ya no es necesario aquí.
        }

        // =================================================================
        // ===== MÉTODO #1: Auditoría de "Tablas de Metrados" =====
        // =================================================================

        public ElementData ProcessSingleElement(ViewSchedule view, string assemblyCode)
        {
            var elementData = new ElementData
            {
                ElementId = view.Id,
                Element = view,
                Nombre = view.Name,
                Categoria = view.Category?.Name ?? "Tabla de Planificación",
                CodigoIdentificacion = assemblyCode
            };

            ScheduleDefinition definition = view.Definition;

            // --- Ejecutar Auditorías ---
            var auditViewName = ProcessViewName(view, assemblyCode);
            var auditFilter = ProcessFilters(definition, elementData.CodigoIdentificacion);
            var auditColumns = ProcessColumns(definition);
            var auditContent = ProcessContent(definition);
            var auditLinks = ProcessIncludeLinks(definition);

            elementData.AuditResults.Add(auditViewName);
            elementData.AuditResults.Add(auditFilter);
            elementData.AuditResults.Add(auditColumns);
            elementData.AuditResults.Add(auditContent);
            elementData.AuditResults.Add(auditLinks);

            // Si el A.C. es inválido, añadir un error fatal
            if (assemblyCode == "INVALID_AC")
            {
                elementData.AuditResults.Add(new AuditItem
                {
                    AuditType = "VIEW NAME",
                    Estado = EstadoParametro.Error,
                    Mensaje = "Error: El nombre no contiene un Assembly Code (ej. C.xx.xx) válido.",
                    IsCorrectable = false
                });
            }

            // --- Almacenar datos para corrección (Tag) ---
            if (auditViewName.Estado == EstadoParametro.Corregir) auditViewName.Tag = auditViewName.ValorCorregido;
            if (auditFilter.Estado == EstadoParametro.Corregir) auditFilter.Tag = auditFilter.Tag;
            if (auditContent.Estado == EstadoParametro.Corregir) auditContent.Tag = true;
            if (auditLinks.Estado == EstadoParametro.Corregir) auditLinks.Tag = true;
            // No se añade Tag para Columnas (Advertencia)

            elementData.DatosCompletos = elementData.AuditResults.All(r => r.Estado == EstadoParametro.Correcto);

            return elementData;
        }

        // =================================================================
        // ===== (NUEVO) Auditoría 0: VIEW NAME (Simplificada) =====
        // =================================================================
        private AuditItem ProcessViewName(ViewSchedule view, string assemblyCode)
        {
            var item = new AuditItem
            {
                AuditType = "VIEW NAME",
                ValorActual = view.Name
            };

            // Caso 1: A.C. no válido (Error)
            if (assemblyCode == "INVALID_AC")
            {
                item.Estado = EstadoParametro.Error;
                item.Mensaje = "Error: El nombre no contiene un Assembly Code (ej. C.xx.xx) válido.";
                item.ValorCorregido = "N/A";
                item.IsCorrectable = false;
                return item;
            }

            string viewName = view.Name;
            bool sufijoCorrecto = viewName.EndsWith(" - RNG", StringComparison.Ordinal);
            string separadorEsperado = " - ";
            bool separadorCorrecto = viewName.Substring(assemblyCode.Length).StartsWith(separadorEsperado);

            if (sufijoCorrecto && separadorCorrecto)
            {
                // Caso 2: Coincide (Correcto)
                item.Estado = EstadoParametro.Correcto;
                item.Mensaje = "Correcto: El formato del nombre de la tabla es válido.";
                item.ValorCorregido = viewName;
                item.IsCorrectable = false;
            }
            else
            {
                // Caso 3: No coincide (Corregir)
                item.Estado = EstadoParametro.Corregir;
                item.Mensaje = "Corregir: El formato del nombre (separador o sufijo) es incorrecto.";
                item.IsCorrectable = true;

                // Construir el nombre corregido
                // 1. Quitar el sufijo incorrecto (si existe)
                string desc = viewName.Substring(assemblyCode.Length);
                if (desc.EndsWith(" - RNG", StringComparison.OrdinalIgnoreCase))
                {
                    desc = desc.Substring(0, desc.Length - " - RNG".Length);
                }
                else if (desc.EndsWith("-RNG", StringComparison.OrdinalIgnoreCase))
                {
                    desc = desc.Substring(0, desc.Length - "-RNG".Length);
                }

                // 2. Limpiar el inicio (quitar guiones o espacios)
                desc = desc.TrimStart(' ', '-');

                // 3. Ensamblar
                item.ValorCorregido = $"{assemblyCode} - {desc} - RNG";
            }

            return item;
        }


        #region Auditoría 1: FILTRO (Lógica de ESTADOS modificada)

        private AuditItem ProcessFilters(ScheduleDefinition definition, string assemblyCode)
        {
            var item = new AuditItem { AuditType = "FILTRO", IsCorrectable = false, Estado = EstadoParametro.Correcto }; // Inicia como Correcto
            var filtrosActuales = definition.GetFilters().ToList();
            var filtrosCorrectosInfo = new List<ScheduleFilterInfo>();

            ScheduleFilter filtroPrincipal = null;
            int indexFiltroPrincipal = -1;
            int contadorFiltrosPrincipales = 0;

            for (int i = 0; i < filtrosActuales.Count; i++)
            {
                var filter = filtrosActuales[i];
                string fieldName = definition.GetField(filter.FieldId).GetName();
                if (IsAssemblyCodeField(fieldName))
                {
                    string valorFiltro = GetFilterValueString(filter);
                    if (valorFiltro.StartsWith("C.", StringComparison.OrdinalIgnoreCase))
                    {
                        contadorFiltrosPrincipales++;
                        if (filtroPrincipal == null) { filtroPrincipal = filter; indexFiltroPrincipal = i; }
                    }
                }
            }

            // (MODIFICADO) ERROR (Fatal)
            if (contadorFiltrosPrincipales > 1)
            {
                item.Estado = EstadoParametro.Error;
                item.Mensaje = "Error: Hay múltiples filtros de Assembly Code con valor que empieza con 'C.'.";
                item.ValorActual = GetFiltersAsString(definition, filtrosActuales);
                item.ValorCorregido = "(No se puede determinar)";
                return item;
            }

            // (MODIFICADO) ERROR (Fatal)
            if (filtroPrincipal == null)
            {
                item.Estado = EstadoParametro.Error;
                item.Mensaje = "Error: No se encontró filtro de Assembly Code con valor que empiece con 'C.'.";
                item.ValorActual = GetFiltersAsString(definition, filtrosActuales);
                item.ValorCorregido = $"Assembly Code Equal {assemblyCode}\nEMPRESA Equal RNG";
                return item;
            }

            // (MODIFICADO) A.C. vs Nombre -> ADVERTENCIA
            string valorPrincipal = GetFilterValueString(filtroPrincipal);
            bool assemblyCodeCorrecto = valorPrincipal.Equals(assemblyCode, StringComparison.OrdinalIgnoreCase);
            if (!assemblyCodeCorrecto)
            {
                item.Estado = EstadoParametro.Vacio; // Advertencia
                item.Mensaje = $"Advertencia: El valor del filtro AC ('{valorPrincipal}') no coincide con el del nombre ('{assemblyCode}').";
            }

            filtrosCorrectosInfo.Add(new ScheduleFilterInfo
            {
                FieldName = definition.GetField(filtroPrincipal.FieldId).GetName(),
                FilterType = filtroPrincipal.FilterType,
                Value = assemblyCode
            });

            ScheduleFilter filtroEmpresa = null;
            int indexFiltroEmpresa = -1;
            for (int i = 0; i < filtrosActuales.Count; i++)
            {
                string fieldName = definition.GetField(filtrosActuales[i].FieldId).GetName();
                if (fieldName.Equals("EMPRESA", StringComparison.OrdinalIgnoreCase))
                {
                    filtroEmpresa = filtrosActuales[i]; indexFiltroEmpresa = i; break;
                }
            }

            // (MODIFICADO) EMPRESA -> CORREGIR
            bool empresaCorrecta = false;
            if (filtroEmpresa != null)
            {
                string valorEmpresa = GetFilterValueString(filtroEmpresa);
                empresaCorrecta = valorEmpresa.Equals("RNG", StringComparison.OrdinalIgnoreCase);

                if (!empresaCorrecta && item.Estado == EstadoParametro.Correcto) // Solo si no hay error/advertencia previa
                {
                    item.Estado = EstadoParametro.Corregir;
                    item.Mensaje = "Corregir: Filtro EMPRESA no es 'RNG'.";
                    item.IsCorrectable = true;
                }
            }
            else if (item.Estado == EstadoParametro.Correcto) // No existe y no hay error/advertencia previa
            {
                item.Estado = EstadoParametro.Corregir;
                item.Mensaje = "Corregir: Falta el filtro EMPRESA = 'RNG'.";
                item.IsCorrectable = true;
            }

            filtrosCorrectosInfo.Add(new ScheduleFilterInfo
            {
                FieldName = "EMPRESA",
                FilterType = ScheduleFilterType.Equal,
                Value = "RNG"
            });

            for (int i = 0; i < filtrosActuales.Count; i++)
            {
                if (i == indexFiltroPrincipal || i == indexFiltroEmpresa) continue;
                var filter = filtrosActuales[i];
                filtrosCorrectosInfo.Add(new ScheduleFilterInfo
                {
                    FieldName = definition.GetField(filter.FieldId).GetName(),
                    FilterType = filter.FilterType,
                    Value = GetFilterValueObject(filter)
                });
            }

            bool ordenCorrecto = true;
            if (filtrosActuales.Count != filtrosCorrectosInfo.Count)
            {
                ordenCorrecto = false;
            }
            else
            {
                for (int i = 0; i < filtrosActuales.Count; i++)
                {
                    var actualInfo = GetFilterInfo(filtrosActuales[i], definition);
                    var correctoInfo = filtrosCorrectosInfo[i];
                    if (actualInfo.FieldName != correctoInfo.FieldName ||
                        actualInfo.FilterType != correctoInfo.FilterType ||
                        !AreFilterValuesEqual(actualInfo.Value, correctoInfo.Value))
                    {
                        ordenCorrecto = false; break;
                    }
                }
            }

            item.ValorActual = GetFiltersAsString(definition, filtrosActuales);
            item.ValorCorregido = string.Join("\n", filtrosCorrectosInfo.Select(f => f.AsString()));

            if (!ordenCorrecto && item.Estado == EstadoParametro.Correcto) // No hay error/advertencia previa
            {
                item.Estado = EstadoParametro.Corregir;
                item.Mensaje = "Corregir: Los filtros no están en el orden correcto o tienen valores diferentes.";
                item.IsCorrectable = true;
            }

            if (item.Estado == EstadoParametro.Correcto)
            {
                item.Mensaje = "Correcto: Filtros correctos.";
            }

            if (item.IsCorrectable) item.Tag = filtrosCorrectosInfo;

            return item;
        }

        // ... (Funciones auxiliares de Filtro: IsAssemblyCodeField, GetFilterInfo, etc. van aquí) ...
        private bool IsAssemblyCodeField(string fieldName)
        {
            if (string.IsNullOrEmpty(fieldName)) return false;
            string fn = fieldName.Trim();
            return fn.Equals("Assembly Code", StringComparison.OrdinalIgnoreCase) ||
                   fn.EndsWith(": Assembly Code", StringComparison.OrdinalIgnoreCase) ||
                   fn.Equals("MATERIAL_ASSEMBLY CODE", StringComparison.OrdinalIgnoreCase) ||
                   fn.EndsWith(": MATERIAL_ASSEMBLY CODE", StringComparison.OrdinalIgnoreCase);
        }
        private ScheduleFilterInfo GetFilterInfo(ScheduleFilter filter, ScheduleDefinition def)
        {
            var field = def.GetField(filter.FieldId);
            return new ScheduleFilterInfo { FieldName = field.GetName(), FilterType = filter.FilterType, Value = GetFilterValueObject(filter) };
        }
        private string GetFiltersAsString(ScheduleDefinition definition, IList<ScheduleFilter> filters)
        {
            if (filters == null || filters.Count == 0) return "(Sin Filtros)";
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
            if (filter.FilterType == ScheduleFilterType.HasValue || filter.FilterType == ScheduleFilterType.HasNoValue) return string.Empty;
            if (filter.IsStringValue) return filter.GetStringValue();
            if (filter.IsDoubleValue) return filter.GetDoubleValue().ToString();
            if (filter.IsElementIdValue) return filter.GetElementIdValue().IntegerValue.ToString();
            return "(valor no legible)";
        }
        private object GetFilterValueObject(ScheduleFilter filter)
        {
            if (filter.FilterType == ScheduleFilterType.HasValue || filter.FilterType == ScheduleFilterType.HasNoValue) return null;
            if (filter.IsStringValue) return filter.GetStringValue();
            if (filter.IsDoubleValue) return filter.GetDoubleValue();
            if (filter.IsElementIdValue) return filter.GetElementIdValue();
            return null;
        }
        private bool AreFilterValuesEqual(object value1, object value2)
        {
            if (value1 == null && value2 == null) return true;
            if (value1 == null || value2 == null) return false;
            if (value1 is string s1 && value2 is string s2) return s1.Equals(s2, StringComparison.OrdinalIgnoreCase);
            if (value1 is ElementId id1 && value2 is ElementId id2) return id1.IntegerValue == id2.IntegerValue;
            return value1.Equals(value2);
        }

        #endregion

        #region Auditoría 2: COLUMNAS (MODIFICADO a Advertencia)

        private AuditItem ProcessColumns(ScheduleDefinition definition)
        {
            var item = new AuditItem
            {
                AuditType = "COLUMNAS",
                IsCorrectable = false // No se puede corregir automáticamente
            };

            var actualHeadings = new List<string>();
            for (int i = 0; i < definition.GetFieldCount(); i++)
            {
                var field = definition.GetField(i);
                if (!field.IsHidden)
                {
                    actualHeadings.Add(field.ColumnHeading.ToUpper().Trim());
                }
            }

            // (MODIFICADO) Conteo -> ADVERTENCIA
            if (actualHeadings.Count != _expectedHeadingCount)
            {
                item.Estado = EstadoParametro.Vacio; // Advertencia
                item.Mensaje = $"Advertencia: Se esperaban {_expectedHeadingCount} columnas visibles, pero se encontraron {actualHeadings.Count}.";
                item.ValorActual = $"Total: {actualHeadings.Count}\n" + string.Join("\n", actualHeadings);
                item.ValorCorregido = $"Total: {_expectedHeadingCount}\n" + string.Join("\n", _expectedHeadings);
                return item;
            }

            bool isCorrect = true;
            // (Lógica de revisión de encabezados omitida por brevedad, es la misma que ya tenías)
            for (int i = 0; i < _expectedHeadingCount; i++)
            {
                string actual = actualHeadings[i];
                string expectedBase = _expectedHeadings[i];
                string corrected = actual;
                if (i == 5)
                {
                    if (actual == expectedBase || actual == "EJES") { }
                    else if (_aliasEncabezados.TryGetValue(actual, out string aliasValue)) corrected = aliasValue;
                    else isCorrect = false;
                }
                else
                {
                    if (actual == expectedBase) { }
                    else if (_aliasEncabezados.TryGetValue(actual, out string aliasValue)) corrected = aliasValue;
                }
                if (i == 5)
                {
                    if (corrected != "AMBIENTE" && corrected != "EJES") isCorrect = false;
                }
                else if (corrected != expectedBase)
                {
                    isCorrect = false;
                }
            }

            item.ValorActual = string.Join("\n", actualHeadings);
            item.ValorCorregido = string.Join("\n", _expectedHeadings);

            // (MODIFICADO) Encabezados -> ADVERTENCIA
            if (!isCorrect)
            {
                item.Estado = EstadoParametro.Vacio; // Advertencia
                item.Mensaje = "Advertencia: Los encabezados no coinciden con el estándar (ALIAS).";
            }
            else
            {
                item.Estado = EstadoParametro.Correcto;
                item.Mensaje = "Correcto: Encabezados correctos.";
            }

            return item;
        }

        #endregion

        #region Auditoría 3: CONTENIDO (Itemize - Prefijo de Mensaje)

        private AuditItem ProcessContent(ScheduleDefinition definition)
        {
            var item = new AuditItem
            {
                AuditType = "CONTENIDO",
                IsCorrectable = true,
                ValorCorregido = "Detallar cada ejemplar: Sí"
            };

            bool isItemized = definition.IsItemized;
            item.ValorActual = $"Detallar cada ejemplar: {(isItemized ? "Sí" : "No")}";

            if (isItemized)
            {
                item.Estado = EstadoParametro.Correcto;
                item.Mensaje = "Correcto: 'Detallar cada ejemplar' está activado.";
                item.IsCorrectable = false;
            }
            else
            {
                item.Estado = EstadoParametro.Corregir;
                item.Mensaje = "Corregir: 'Detallar cada ejemplar' está desactivado.";
                item.Tag = true;
            }
            return item;
        }

        #endregion

        #region (NUEVO) Auditoría 4: INCLUDE LINKS

        private AuditItem ProcessIncludeLinks(ScheduleDefinition definition)
        {
            var item = new AuditItem
            {
                AuditType = "LINKS",
                IsCorrectable = true,
                ValorCorregido = "Incluir elementos de links: Sí"
            };

            bool includeLinks = definition.IncludeLinkedFiles;
            item.ValorActual = $"Incluir elementos de links: {(includeLinks ? "Sí" : "No")}";

            if (includeLinks)
            {
                item.Estado = EstadoParametro.Correcto;
                item.Mensaje = "Correcto: 'Include elements in links' está activado.";
                item.IsCorrectable = false;
            }
            else
            {
                item.Estado = EstadoParametro.Corregir;
                item.Mensaje = "Corregir: 'Include elements in links' está desactivado.";
                item.Tag = true; // Marcar para corrección
            }
            return item;
        }

        #endregion

        // =================================================================
        // ===== MÉTODOS DE CREACIÓN DE TRABAJOS (Jobs) =====
        // =================================================================

        public ElementData CreateRenamingJob(ViewSchedule view)
        {
            string nombreActual = view.Name;
            string nombreCorregido = nombreActual.Replace("C.", "SOPORTE.");
            var jobData = new RenamingJobData { NuevoNombre = nombreCorregido, NuevoGrupoVista = "00 TRABAJO EN PROCESO - WIP", NuevoSubGrupoVista = "SOPORTE DE METRADOS" };
            return CreateJobElementData(view, "CLASIFICACIÓN (SOPORTE)", "Corregir: Tabla de soporte mal clasificada.", jobData);
        }

        public ElementData CreateManualRenamingJob(ViewSchedule view)
        {
            var jobData = new RenamingJobData { NuevoNombre = view.Name, NuevoGrupoVista = "REVISAR", NuevoSubGrupoVista = "METRADO MANUAL" };
            return CreateJobElementData(view, "MANUAL", "Corregir: Tabla de Metrado Manual. Se reclasificará.", jobData);
        }

        // (NUEVO)
        public ElementData CreateCopyReclassifyJob(ViewSchedule view)
        {
            var jobData = new RenamingJobData { NuevoNombre = view.Name, NuevoGrupoVista = "REVISAR", NuevoSubGrupoVista = string.Empty };
            return CreateJobElementData(view, "COPIA", "Corregir: La tabla parece ser una copia y será reclasificada.", jobData);
        }

        // (NUEVO)
        public ElementData CreateWipReclassifyJob(ViewSchedule view)
        {
            var jobData = new RenamingJobData { NuevoNombre = view.Name, NuevoGrupoVista = "00 TRABAJO EN PROCESO - WIP", NuevoSubGrupoVista = "SOPORTE BIM" };
            return CreateJobElementData(view, "CLASIFICACIÓN (WIP)", "Corregir: Tabla de trabajo interno. Será reclasificada.", jobData);
        }

        // (NUEVO) Helper para crear trabajos
        private ElementData CreateJobElementData(ViewSchedule view, string auditType, string mensaje, RenamingJobData jobData)
        {
            // (NUEVO) Leer valores actuales
            string grupoActual = view.LookupParameter("GRUPO DE VISTA")?.AsString() ?? "(Vacío)";
            string subGrupoActual = view.LookupParameter("SUBGRUPO DE VISTA")?.AsString() ?? "(Vacío)";

            // (NUEVO) Construir strings de valor
            string valorActualStr = $"Grupo: {grupoActual}\nSubGrupo: {subGrupoActual}";
            string valorCorregidoStr = $"Grupo: {jobData.NuevoGrupoVista}\nSubGrupo: {jobData.NuevoSubGrupoVista}";

            var elementData = new ElementData
            {
                ElementId = view.Id,
                Element = view,
                Nombre = view.Name,
                Categoria = view.Category?.Name ?? "Tabla de Planificación",
                CodigoIdentificacion = "N/A",
                DatosCompletos = false
            };

            var auditItem = new AuditItem
            {
                AuditType = auditType,
                IsCorrectable = true,
                Estado = EstadoParametro.Corregir,
                Mensaje = mensaje,
                ValorActual = valorActualStr,
                ValorCorregido = valorCorregidoStr,
                Tag = jobData
            };
            elementData.AuditResults.Add(auditItem);
            return elementData;
        }
    }
}