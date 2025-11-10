// Plugins/Tablas/ScheduleProcessor.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using TL60_RevisionDeTablas.Models;
using TL60_RevisionDeTablas.Core;

namespace TL60_RevisionDeTablas.Plugins.Tablas
{
    public class ScheduleProcessor
    {
        // ==========================================================
        // ===== (¡CORRECCIÓN!) CAMPOS DE CLASE REINSERTADOS
        // ==========================================================
        private readonly Document _doc;
        private readonly GoogleSheetsService _sheetsService;
        private readonly UniclassDataService _uniclassService;
        private readonly string _docTitle;
        private readonly string _spreadsheetId;
        private readonly List<string> _ejesKeywords;

        // (Campos estáticos)
        private static readonly Regex _acRegex = new Regex(@"^C\.(\d{2,3}\.)+\d{2,3}");
        private static readonly List<string> NOMBRES_WIP = new List<string>
        {
            "TL", "TITO", "PDONTADENEA", "ANDREA", "EFRAIN",
            "PROYECTOSBIM", "ASISTENTEBIM", "LUIS", "DIEGO", "JORGE", "MIGUEL"
        };
        private static readonly List<string> MODELOS_ARQUITECTURA = new List<string>
        {
            "200114-CCC02-MO-AR-045500", "200114-CCC02-MO-AR-045600",
            "200114-CCC02-MO-AR-046900", "200114-CCC02-MO-AR-047000",
            "200114-CCC02-MO-AR-047100", "200114-CCC02-MO-AR-047200",
            "200114-CCC02-MO-AR-047300", "200114-CCC02-MO-AR-047400",
            "200114-CCC02-MO-AR-047500", "200114-CCC02-MO-AR-047600",
            "200114-CCC02-MO-AR-047700", "200114-CCC02-MO-AR-047800",
            "200114-CCC02-MO-AR-047900", "200114-CCC02-MO-AR-048000",
            "200114-CCC02-MO-AR-000410"
        };
        private static readonly List<string> MODELOS_ESTRUCTURAS = new List<string>
        {
            "200114-CCC02-MO-ES-045500", "200114-CCC02-MO-ES-045600",
            "200114-CCC02-MO-ES-046900", "200114-CCC02-MO-ES-047000",
            "200114-CCC02-MO-ES-047100", "200114-CCC02-MO-ES-047200",
            "200114-CCC02-MO-ES-047300", "200114-CCC02-MO-ES-047400",
            "200114-CCC02-MO-ES-047500", "200114-CCC02-MO-ES-047600",
            "200114-CCC02-MO-ES-047700", "200114-CCC02-MO-ES-047800",
            "200114-CCC02-MO-ES-047900", "200114-CCC02-MO-ES-048000",
            "200114-CCC02-MO-ES-000410"
        };

        // (¡CORRECCIÓN! Este campo faltaba)
        private readonly Dictionary<string, string> _aliasEncabezados = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "DIGO", "CODIGO" }, { "DESCRIP", "DESCRIPCION" }, { "ACTIV", "ACTIVO" },
            { "DULO", "MODULO" }, { "NIVE", "NIVEL" }, { "AMBIEN", "AMBIENTE" },
            { "EJE", "EJES" }, { "PARC", "PARCIAL" }, { "UNID", "UNIDAD" },
            { "ID DE ELEMENTO", "ID" },
            { "ID", "ID" }
        };

        // (¡CORRECCIÓN! Este campo faltaba)
        private readonly List<string> _baseExpectedHeadings = new List<string>
        {
            "CODIGO", "DESCRIPCION", "ACTIVO", "MODULO", "NIVEL",
            "AMBIENTE", // <--- Este es el índice 5
            "PARCIAL", "UNIDAD", "ID"
        };
        private const int _expectedHeadingCount = 9;


        // (Constructor de 4 argumentos, al que MainCommand llama)
        public ScheduleProcessor(Document doc, GoogleSheetsService sheetsService, UniclassDataService uniclassService, string spreadsheetId)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _sheetsService = sheetsService ?? throw new ArgumentNullException(nameof(sheetsService));
            _uniclassService = uniclassService ?? throw new ArgumentNullException(nameof(uniclassService));
            _docTitle = doc.Title;
            _spreadsheetId = spreadsheetId;
            _ejesKeywords = new List<string>();

            LoadEjesKeywords();
        }

        private void LoadEjesKeywords()
        {
            try
            {
                var data = _sheetsService.ReadData(_spreadsheetId, "'FIELD EJES'!B2:B");
                if (data == null || data.Count == 0) return;

                foreach (var row in data)
                {
                    string keyword = GoogleSheetsService.GetCellValue(row, 0);
                    if (!string.IsNullOrWhiteSpace(keyword))
                    {
                        _ejesKeywords.Add(keyword.ToUpper());
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al cargar 'FIELD EJES': {ex.Message}");
            }
        }

        // (Este método es PÚBLICO y SÍ existe)
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
            var auditColumns = ProcessColumns(view);
            var auditContent = ProcessContent(definition);
            var auditLinks = ProcessIncludeLinks(definition);
            var auditParcial = ProcessParcialFormat(definition);

            elementData.AuditResults.Add(auditViewName);
            elementData.AuditResults.Add(auditFilter);
            elementData.AuditResults.Add(auditColumns);
            elementData.AuditResults.Add(auditContent);
            elementData.AuditResults.Add(auditLinks);
            elementData.AuditResults.Add(auditParcial);

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
            if (auditColumns.Estado == EstadoParametro.Corregir) auditColumns.Tag = auditColumns.Tag;
            if (auditParcial.Estado == EstadoParametro.Corregir) auditParcial.Tag = auditParcial.Tag;

            elementData.DatosCompletos = elementData.AuditResults.All(r => r.Estado == EstadoParametro.Correcto);

            return elementData;
        }

        // ... (Todos los métodos de auditoría privados: ProcessViewName, ProcessFilters, etc. van aquí) ...

        #region Auditoría 1: VIEW NAME
        private AuditItem ProcessViewName(ViewSchedule view, string assemblyCode)
        {
            var item = new AuditItem
            {
                AuditType = "VIEW NAME",
                ValorActual = view.Name
            };
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
                item.Estado = EstadoParametro.Correcto;
                item.Mensaje = "Correcto: El formato del nombre de la tabla es válido.";
                item.ValorCorregido = viewName;
                item.IsCorrectable = false;
            }
            else
            {
                item.Estado = EstadoParametro.Corregir;
                item.Mensaje = "Corregir: El formato del nombre (separador o sufijo) es incorrecto.";
                item.IsCorrectable = true;
                string desc = viewName.Substring(assemblyCode.Length);
                if (desc.EndsWith(" - RNG", StringComparison.OrdinalIgnoreCase))
                {
                    desc = desc.Substring(0, desc.Length - " - RNG".Length);
                }
                else if (desc.EndsWith("-RNG", StringComparison.OrdinalIgnoreCase))
                {
                    desc = desc.Substring(0, desc.Length - "-RNG".Length);
                }
                desc = desc.TrimStart(' ', '-');
                item.ValorCorregido = $"{assemblyCode} - {desc} - RNG";
            }
            return item;
        }
        #endregion

        #region Auditoría 2: FILTRO
        private AuditItem ProcessFilters(ScheduleDefinition definition, string assemblyCode)
        {
            var item = new AuditItem { AuditType = "FILTRO", IsCorrectable = false, Estado = EstadoParametro.Correcto };
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
            if (contadorFiltrosPrincipales > 1)
            {
                item.Estado = EstadoParametro.Error;
                item.Mensaje = "Error: Hay múltiples filtros de Assembly Code con valor que empieza con 'C.'.";
                item.ValorActual = GetFiltersAsString(definition, filtrosActuales);
                item.ValorCorregido = "(No se puede determinar)";
                return item;
            }
            if (filtroPrincipal == null)
            {
                item.Estado = EstadoParametro.Error;
                item.Mensaje = "Error: No se encontró filtro de Assembly Code con valor que empiece con 'C.'.";
                item.ValorActual = GetFiltersAsString(definition, filtrosActuales);
                item.ValorCorregido = $"Assembly Code Equal {assemblyCode}\nEMPRESA Equal RNG";
                return item;
            }
            string valorPrincipal = GetFilterValueString(filtroPrincipal);
            bool assemblyCodeCorrecto = valorPrincipal.Equals(assemblyCode, StringComparison.OrdinalIgnoreCase);
            if (!assemblyCodeCorrecto)
            {
                item.Estado = EstadoParametro.Vacio;
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
            bool empresaCorrecta = false;
            if (filtroEmpresa != null)
            {
                string valorEmpresa = GetFilterValueString(filtroEmpresa);
                empresaCorrecta = valorEmpresa.Equals("RNG", StringComparison.OrdinalIgnoreCase);
                if (!empresaCorrecta && item.Estado == EstadoParametro.Correcto)
                {
                    item.Estado = EstadoParametro.Corregir;
                    item.Mensaje = "Corregir: Filtro EMPRESA no es 'RNG'.";
                    item.IsCorrectable = true;
                }
            }
            else if (item.Estado == EstadoParametro.Correcto)
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
            if (!ordenCorrecto && item.Estado == EstadoParametro.Correcto)
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

        #region Auditoría 3: COLUMNAS
        private string GetDynamicExpectedHeader(string viewName)
        {
            bool isEstriModel = MODELOS_ESTRUCTURAS.Any(modelName => _docTitle.Contains(modelName));
            if (isEstriModel)
            {
                return "EJES";
            }
            bool isArchiModel = MODELOS_ARQUITECTURA.Any(modelName => _docTitle.Contains(modelName));
            if (isArchiModel)
            {
                string viewNameUpper = viewName.ToUpper();
                bool usesEjes = _ejesKeywords.Any(keyword => viewNameUpper.Contains(keyword));
                return usesEjes ? "EJES" : "AMBIENTE";
            }
            return "AMBIENTE";
        }
        private AuditItem ProcessColumns(ViewSchedule view)
        {
            var definition = view.Definition;
            var item = new AuditItem
            {
                AuditType = "COLUMNAS",
                IsCorrectable = false,
                Estado = EstadoParametro.Correcto
            };
            var dynamicExpectedHeadings = new List<string>(_baseExpectedHeadings);
            dynamicExpectedHeadings[5] = GetDynamicExpectedHeader(view.Name);
            var actualHeadings = new List<string>();
            var actualFields = new List<ScheduleField>();
            for (int i = 0; i < definition.GetFieldCount(); i++)
            {
                var field = definition.GetField(i);
                if (!field.IsHidden)
                {
                    actualHeadings.Add(field.ColumnHeading.ToUpper().Trim());
                    actualFields.Add(field);
                }
            }
            if (actualHeadings.Count > _expectedHeadingCount)
            {
                item.Estado = EstadoParametro.Corregir;
                item.IsCorrectable = true;
                var fieldsToHide = new List<ScheduleField>();
                var fieldsToHideNames = new List<string>();
                for (int i = 0; i < actualFields.Count; i++)
                {
                    if (!dynamicExpectedHeadings.Contains(actualHeadings[i], StringComparer.OrdinalIgnoreCase))
                    {
                        fieldsToHide.Add(actualFields[i]);
                        fieldsToHideNames.Add(actualFields[i].ColumnHeading);
                    }
                }
                item.Tag = fieldsToHide;
                item.ValorActual = string.Join("\n", actualHeadings);
                item.ValorCorregido = string.Join("\n", dynamicExpectedHeadings);
                item.Mensaje = $"A corregir: Se esperaban 9 columnas visibles, pero se encontraron {actualHeadings.Count}.\n" +
                               $"Se ocultarán los siguientes campos:\n" +
                               string.Join("\n", fieldsToHideNames);
                return item;
            }
            if (actualHeadings.Count == _expectedHeadingCount)
            {
                var headingsToFix = new Dictionary<ScheduleField, string>();
                var correctedHeadings = new List<string>();
                for (int i = 0; i < _expectedHeadingCount; i++)
                {
                    string actual = actualHeadings[i];
                    string expected = dynamicExpectedHeadings[i];
                    ScheduleField field = actualFields[i];
                    if (actual.Equals(expected, StringComparison.OrdinalIgnoreCase))
                    {
                        correctedHeadings.Add(actual);
                        continue;
                    }
                    item.Estado = EstadoParametro.Corregir;
                    item.IsCorrectable = true;
                    string correctedValue = expected;
                    bool foundAlias = false;
                    foreach (var aliasKey in _aliasEncabezados.Keys.OrderByDescending(k => k.Length))
                    {
                        if (actual.Contains(aliasKey))
                        {
                            correctedValue = _aliasEncabezados[aliasKey];
                            foundAlias = true;
                            break;
                        }
                    }
                    if (foundAlias && !correctedValue.Equals(expected, StringComparison.OrdinalIgnoreCase))
                    {
                        correctedValue = expected;
                    }
                    correctedHeadings.Add(correctedValue);
                    headingsToFix[field] = correctedValue;
                }
                item.ValorActual = string.Join("\n", actualHeadings);
                item.ValorCorregido = string.Join("\n", correctedHeadings);
                if (item.IsCorrectable)
                {
                    item.Mensaje = "Corregir: Los encabezados no coinciden con el estándar (ALIAS).";
                    item.Tag = headingsToFix;
                }
                else
                {
                    item.Estado = EstadoParametro.Correcto;
                    item.Mensaje = "Correcto: Encabezados correctos.";
                }
                return item;
            }
            item.Estado = EstadoParametro.Vacio;
            item.Mensaje = $"Advertencia: Se esperaban 9 columnas visibles, pero se encontraron {actualHeadings.Count}.";
            item.ValorActual = $"Total: {actualHeadings.Count}\n" + string.Join("\n", actualHeadings);
            item.ValorCorregido = $"Total: {_expectedHeadingCount}\n" + string.Join("\n", dynamicExpectedHeadings);
            return item;
        }
        #endregion

        #region Auditoría 4: CONTENIDO
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

        #region Auditoría 5: LINKS
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
                item.Tag = true;
            }
            return item;
        }
        #endregion

        #region Auditoría 6: FORMATO "PARCIAL" (Corregido)
        private AuditItem ProcessParcialFormat(ScheduleDefinition definition)
        {
            var item = new AuditItem
            {
                AuditType = "FORMATO PARCIAL",
                Estado = EstadoParametro.Correcto,
                IsCorrectable = false
            };
            ScheduleField parcialField = null;
            var parcialAliases = _aliasEncabezados
                .Where(kvp => kvp.Value.Equals("PARCIAL", StringComparison.OrdinalIgnoreCase))
                .Select(kvp => kvp.Key)
                .ToList();
            parcialAliases.Add("PARCIAL");
            for (int i = 0; i < definition.GetFieldCount(); i++)
            {
                var field = definition.GetField(i);
                string heading = field.ColumnHeading.ToUpper().Trim();
                if (heading.Equals("PARCIAL", StringComparison.OrdinalIgnoreCase) ||
                    parcialAliases.Any(alias => heading.Contains(alias)))
                {
                    parcialField = field;
                    break;
                }
            }

            if (parcialField == null)
            {
                item.Mensaje = "N/A (No se encontró la columna 'PARCIAL')";
                return item;
            }

            try
            {
                FormatOptions options = parcialField.GetFormatOptions();

                ForgeTypeId symbolId = options.GetSymbolTypeId();
                bool isCorrect = !options.UseDefault && (symbolId == null || string.IsNullOrEmpty(symbolId.TypeId));

                item.ValorActual = $"Símbolo: {symbolId?.TypeId ?? "NINGUNO"} (Usa Default: {options.UseDefault})";
                item.ValorCorregido = $"Símbolo: NINGUNO (Usa Default: False)";

                if (isCorrect)
                {
                    item.Mensaje = "Correcto: El formato de 'PARCIAL' no muestra unidad.";
                }
                else
                {
                    item.Estado = EstadoParametro.Corregir;
                    item.Mensaje = "Corregir: La columna 'PARCIAL' debe ocultar el símbolo de unidad.";
                    item.IsCorrectable = true;
                    item.Tag = parcialField.FieldId;
                }
            }
            catch (Exception ex)
            {
                item.Estado = EstadoParametro.Error;
                item.Mensaje = $"Error al leer formato de 'PARCIAL': {ex.Message}";
            }

            return item;
        }
        #endregion

        #region Creación de Jobs (Modificado)
        public ElementData CreateRenamingJob(ViewSchedule view)
        {
            string nombreActual = view.Name;
            string nombreCorregido = nombreActual.Replace("C.", "SOPORTE.");
            var jobData = new RenamingJobData { NuevoNombre = nombreCorregido, NuevoGrupoVista = "00 TRABAJO EN PROCESO - WIP", NuevoSubGrupoVista = "SOPORTE DE METRADOS" };
            return CreateJobElementData(view, "CLASIFICACIÓN (SOPORTE)", "Corregir: Tabla de soporte mal clasificada.", jobData);
        }
        public ElementData CreateManualRenamingJob(ViewSchedule view)
        {
            string nombreCorregido = view.Name.Replace("C.", "SOPORTE.");
            var jobData = new RenamingJobData { NuevoNombre = nombreCorregido, NuevoGrupoVista = "REVISAR", NuevoSubGrupoVista = "METRADO MANUAL" };
            return CreateJobElementData(view, "MANUAL", "Corregir: Tabla de Metrado Manual. Se reclasificará y renombrará.", jobData);
        }
        public ElementData CreateCopyReclassifyJob(ViewSchedule view)
        {
            string nombreCorregido = view.Name.Replace("C.", "SOPORTE.");
            var jobData = new RenamingJobData { NuevoNombre = nombreCorregido, NuevoGrupoVista = "REVISAR", NuevoSubGrupoVista = string.Empty };
            return CreateJobElementData(view, "COPIA", "Corregir: La tabla parece ser una copia. Se reclasificará y renombrará.", jobData);
        }
        public ElementData CreateWipReclassifyJob(ViewSchedule view)
        {
            string nombreCorregido = view.Name.Replace("C.", "SOPORTE.");
            var jobData = new RenamingJobData { NuevoNombre = nombreCorregido, NuevoGrupoVista = "00 TRABAJO EN PROCESO - WIP", NuevoSubGrupoVista = "SOPORTE BIM" };
            return CreateJobElementData(view, "CLASIFICACIÓN (WIP)", "Corregir: Tabla de trabajo interno. Será reclasificada y renombrada.", jobData);
        }
        private ElementData CreateJobElementData(ViewSchedule view, string auditType, string mensaje, RenamingJobData jobData)
        {
            string grupoActual = view.LookupParameter("GRUPO DE VISTA")?.AsString() ?? "(Vacío)";
            string subGrupoActual = view.LookupParameter("SUBGRUPO DE VISTA")?.AsString() ?? "(Vacío)";
            string valorActualStr = $"NOMBRE: {view.Name}\nGRUPO: {grupoActual}\nSUBGRUPO: {subGrupoActual}";
            string valorCorregidoStr = $"NOMBRE: {jobData.NuevoNombre}\nGRUPO: {jobData.NuevoGrupoVista}\nSUBGRUPO: {jobData.NuevoSubGrupoVista}";
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
        #endregion
    }
}