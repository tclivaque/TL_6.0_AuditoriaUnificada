// Plugins/Uniclass/Services/UniclassAuditService.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TL60_RevisionDeTablas.Core;
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Plugins.Uniclass.Services
{
    /// <summary>
    /// Servicio de auditoría de parámetros Uniclass
    /// </summary>
    public class UniclassAuditService
    {
        private readonly Document _doc;
        private readonly UniclassDataService _uniclassService;
        private readonly List<BuiltInCategory> _categoriasAuditar;
        private readonly StringBuilder _debugLog;
        private const int DEBUG_ELEMENT_ID = 5399549;

        private static readonly string[] UNICLASS_PARAMETERS = new[]
        {
            "Classification.Uniclass.EF.Number",
            "Classification.Uniclass.EF.Description",
            "Classification.Uniclass.Pr.Number",
            "Classification.Uniclass.Pr.Description",
            "Classification.Uniclass.Ss.Number",
            "Classification.Uniclass.Ss.Description"
        };

        public UniclassAuditService(
            Document doc,
            UniclassDataService uniclassService,
            List<BuiltInCategory> categoriasAuditar)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _uniclassService = uniclassService ?? throw new ArgumentNullException(nameof(uniclassService));
            _categoriasAuditar = categoriasAuditar ?? throw new ArgumentNullException(nameof(categoriasAuditar));
            _debugLog = new StringBuilder();
        }

        /// <summary>
        /// Procesa todos los elementos del proyecto y genera datos de auditoría
        /// </summary>
        public List<ElementData> ProcessElements()
        {
            var elementosData = new List<ElementData>();

            _debugLog.AppendLine("=".PadRight(80, '='));
            _debugLog.AppendLine("DEBUG UNICLASS AUDIT SERVICE");
            _debugLog.AppendLine($"Documento: {_doc.Title}");
            _debugLog.AppendLine($"Fecha: {DateTime.Now}");
            _debugLog.AppendLine($"Buscando elemento con ID: {DEBUG_ELEMENT_ID}");
            _debugLog.AppendLine("=".PadRight(80, '='));
            _debugLog.AppendLine();

            // Crear filtro multi-categoría
            var categoryFilters = _categoriasAuditar
                .Select(cat => new ElementCategoryFilter(cat) as ElementFilter)
                .ToList();

            var categoryFilter = new LogicalOrFilter(categoryFilters);

            // Obtener todos los elementos de las categorías a auditar
            var elementos = new FilteredElementCollector(_doc)
                .WherePasses(categoryFilter)
                .WhereElementIsNotElementType()
                .ToElements();

            // Buscar si existe el elemento con ID específico (puede ser instancia o tipo)
            Element debugElement = null;
            try
            {
                debugElement = _doc.GetElement(new ElementId(DEBUG_ELEMENT_ID));
                if (debugElement != null)
                {
                    _debugLog.AppendLine($"Elemento encontrado con ID {DEBUG_ELEMENT_ID}:");
                    _debugLog.AppendLine($"  - Nombre: {debugElement.Name}");
                    _debugLog.AppendLine($"  - Categoría: {debugElement.Category?.Name ?? "N/A"}");
                    _debugLog.AppendLine($"  - Es ElementType: {debugElement is ElementType}");

                    if (debugElement is ElementType)
                    {
                        _debugLog.AppendLine($"  - Es un TIPO de elemento");
                    }
                    else
                    {
                        ElementId typeId = debugElement.GetTypeId();
                        _debugLog.AppendLine($"  - Es una INSTANCIA, su tipo tiene ID: {typeId?.IntegerValue ?? -1}");
                        if (typeId != null && typeId != ElementId.InvalidElementId)
                        {
                            ElementType debugTipo = _doc.GetElement(typeId) as ElementType;
                            if (debugTipo != null)
                            {
                                _debugLog.AppendLine($"  - Nombre del tipo: {debugTipo.Name}");
                            }
                        }
                    }
                    _debugLog.AppendLine();
                }
                else
                {
                    _debugLog.AppendLine($"❌ NO se encontró ningún elemento con ID {DEBUG_ELEMENT_ID} en el proyecto");
                    _debugLog.AppendLine();
                }
            }
            catch (Exception ex)
            {
                _debugLog.AppendLine($"Error al buscar elemento {DEBUG_ELEMENT_ID}: {ex.Message}");
                _debugLog.AppendLine();
            }

            // Agrupar elementos por tipo para procesar por tipo
            var elementosPorTipo = elementos
                .Where(e => e.GetTypeId() != null && e.GetTypeId() != ElementId.InvalidElementId)
                .GroupBy(e => e.GetTypeId());

            _debugLog.AppendLine($"Total de tipos a procesar: {elementosPorTipo.Count()}");
            _debugLog.AppendLine();

            foreach (var grupoTipo in elementosPorTipo)
            {
                ElementType tipo = _doc.GetElement(grupoTipo.Key) as ElementType;
                if (tipo == null) continue;

                // Obtener Assembly Code del tipo usando la clase compartida
                string assemblyCode = AssemblyCodeExtractor.GetAssemblyCode(grupoTipo.First(), _doc);

                // Procesar este tipo
                var elementData = ProcessElementType(tipo, assemblyCode, grupoTipo.Count());
                if (elementData != null)
                {
                    elementosData.Add(elementData);
                }
            }

            // Guardar archivo de debug
            SaveDebugLog();

            return elementosData;
        }

        /// <summary>
        /// Procesa un tipo de elemento específico
        /// </summary>
        private ElementData ProcessElementType(ElementType tipo, string assemblyCode, int instanciasCount)
        {
            // Debug logging para el elemento específico
            bool isDebugElement = tipo.Id.IntegerValue == DEBUG_ELEMENT_ID;

            if (isDebugElement)
            {
                _debugLog.AppendLine($">>> ELEMENTO ENCONTRADO: ID = {tipo.Id.IntegerValue}");
                _debugLog.AppendLine($"    Nombre: {tipo.Name}");
                _debugLog.AppendLine($"    Categoría: {tipo.Category?.Name ?? "N/A"}");
                _debugLog.AppendLine($"    Assembly Code extraído: '{assemblyCode}'");
                _debugLog.AppendLine($"    Assembly Code es válido (empieza con 'C.'): {AssemblyCodeExtractor.IsValidAssemblyCode(assemblyCode)}");
                _debugLog.AppendLine();
            }

            var elementData = new ElementData
            {
                ElementId = tipo.Id,
                Element = tipo,
                Nombre = tipo.Name,
                Categoria = "UNICLASS",
                CodigoIdentificacion = assemblyCode,
                DatosCompletos = true
            };

            // Determinar valores correctos según Assembly Code
            Dictionary<string, string> valoresCorrectos = GetValoresCorrectos(assemblyCode, isDebugElement);

            // Auditar cada uno de los 6 parámetros
            foreach (var paramName in UNICLASS_PARAMETERS)
            {
                AuditarParametro(tipo, paramName, assemblyCode, valoresCorrectos, elementData, isDebugElement);
            }

            if (isDebugElement)
            {
                _debugLog.AppendLine();
                _debugLog.AppendLine("    RESUMEN DE AUDITORÍA:");
                _debugLog.AppendLine($"    - Parámetros correctos: {elementData.ParametrosCorrectos.Count}");
                _debugLog.AppendLine($"    - Parámetros a actualizar: {elementData.ParametrosActualizar.Count}");
                _debugLog.AppendLine($"    - Mensajes de error: {elementData.Mensajes.Count}");
                _debugLog.AppendLine();
            }

            return elementData;
        }

        /// <summary>
        /// Determina los valores correctos para los parámetros Uniclass según el Assembly Code
        /// </summary>
        private Dictionary<string, string> GetValoresCorrectos(string assemblyCode, bool isDebugElement = false)
        {
            var valoresCorrectos = new Dictionary<string, string>();

            if (isDebugElement)
            {
                _debugLog.AppendLine("    --- DETERMINANDO VALORES CORRECTOS ---");
            }

            // Si el Assembly Code NO empieza con "C.", llenar con "-"
            if (!AssemblyCodeExtractor.IsValidAssemblyCode(assemblyCode))
            {
                if (isDebugElement)
                {
                    _debugLog.AppendLine($"    Assembly Code NO es válido (no empieza con 'C.')");
                    _debugLog.AppendLine($"    Todos los parámetros se llenarán con '-'");
                }

                foreach (var paramName in UNICLASS_PARAMETERS)
                {
                    valoresCorrectos[paramName] = "-";
                }
                return valoresCorrectos;
            }

            if (isDebugElement)
            {
                _debugLog.AppendLine($"    Assembly Code ES VÁLIDO (empieza con 'C.')");
                _debugLog.AppendLine($"    Buscando '{assemblyCode}' en matriz Uniclass...");
            }

            // Assembly Code válido (empieza con "C.")
            // Buscar en la matriz Uniclass
            var uniclassData = _uniclassService.GetUniclassParameters(assemblyCode);

            if (uniclassData == null)
            {
                if (isDebugElement)
                {
                    _debugLog.AppendLine($"    ❌ NO SE ENCONTRÓ '{assemblyCode}' en la matriz Uniclass");
                    _debugLog.AppendLine($"    Se marcará como ERROR");
                }

                // No se encontró en la matriz, marcar error
                foreach (var paramName in UNICLASS_PARAMETERS)
                {
                    valoresCorrectos[paramName] = $"ERROR: {assemblyCode} no encontrado en matriz";
                }
                return valoresCorrectos;
            }

            if (isDebugElement)
            {
                _debugLog.AppendLine($"    ✓ SE ENCONTRÓ '{assemblyCode}' en la matriz Uniclass");
                _debugLog.AppendLine($"    ShouldAudit (UNICLAS = '✓'): {uniclassData.ShouldAudit}");
                _debugLog.AppendLine();
                _debugLog.AppendLine("    VALORES EN LA MATRIZ:");
                _debugLog.AppendLine($"      EF.Number: '{uniclassData.EF_Number ?? "(vacío)"}'");
                _debugLog.AppendLine($"      EF.Description: '{uniclassData.EF_Description ?? "(vacío)"}'");
                _debugLog.AppendLine($"      Pr.Number: '{uniclassData.Pr_Number ?? "(vacío)"}'");
                _debugLog.AppendLine($"      Pr.Description: '{uniclassData.Pr_Description ?? "(vacío)"}'");
                _debugLog.AppendLine($"      Ss.Number: '{uniclassData.Ss_Number ?? "(vacío)"}'");
                _debugLog.AppendLine($"      Ss.Description: '{uniclassData.Ss_Description ?? "(vacío)"}'");
                _debugLog.AppendLine();
            }

            // Verificar columna UNICLAS
            if (!uniclassData.ShouldAudit)
            {
                if (isDebugElement)
                {
                    _debugLog.AppendLine($"    UNICLAS NO tiene '✓' → Todos los parámetros se llenarán con '-'");
                }

                // UNICLAS no tiene "✓", llenar con "-"
                foreach (var paramName in UNICLASS_PARAMETERS)
                {
                    valoresCorrectos[paramName] = "-";
                }
                return valoresCorrectos;
            }

            if (isDebugElement)
            {
                _debugLog.AppendLine($"    UNICLAS tiene '✓' → Se usarán los valores de la matriz");
            }

            // UNICLAS tiene "✓", usar valores de la matriz
            valoresCorrectos["Classification.Uniclass.EF.Number"] = uniclassData.EF_Number ?? string.Empty;
            valoresCorrectos["Classification.Uniclass.EF.Description"] = uniclassData.EF_Description ?? string.Empty;
            valoresCorrectos["Classification.Uniclass.Pr.Number"] = uniclassData.Pr_Number ?? string.Empty;
            valoresCorrectos["Classification.Uniclass.Pr.Description"] = uniclassData.Pr_Description ?? string.Empty;
            valoresCorrectos["Classification.Uniclass.Ss.Number"] = uniclassData.Ss_Number ?? string.Empty;
            valoresCorrectos["Classification.Uniclass.Ss.Description"] = uniclassData.Ss_Description ?? string.Empty;

            return valoresCorrectos;
        }

        /// <summary>
        /// Audita un parámetro individual
        /// </summary>
        private void AuditarParametro(
            ElementType tipo,
            string paramName,
            string assemblyCode,
            Dictionary<string, string> valoresCorrectos,
            ElementData elementData,
            bool isDebugElement = false)
        {
            // Obtener valor actual del parámetro
            Parameter param = tipo.LookupParameter(paramName);

            if (param == null)
            {
                // Parámetro no existe en el tipo
                elementData.DatosCompletos = false;
                elementData.Mensajes.Add($"Error: El tipo '{tipo.Name}' no tiene el parámetro '{paramName}'");

                if (isDebugElement)
                {
                    _debugLog.AppendLine($"    ⚠ Parámetro '{paramName}' NO EXISTE en el tipo");
                }
                return;
            }

            string valorActual = param.AsString() ?? string.Empty;
            string valorCorrecto = valoresCorrectos.ContainsKey(paramName) ? valoresCorrectos[paramName] : string.Empty;

            if (isDebugElement)
            {
                _debugLog.AppendLine($"    Parámetro: {paramName}");
                _debugLog.AppendLine($"      Valor actual: '{valorActual}'");
                _debugLog.AppendLine($"      Valor correcto: '{valorCorrecto}'");
            }

            // Caso especial: valor correcto es un mensaje de error
            if (valorCorrecto.StartsWith("ERROR:"))
            {
                elementData.DatosCompletos = false;
                elementData.Mensajes.Add(valorCorrecto);

                if (isDebugElement)
                {
                    _debugLog.AppendLine($"      Estado: ERROR");
                }
                return;
            }

            // Comparar valores
            if (valorActual == valorCorrecto)
            {
                // Parámetro correcto
                elementData.ParametrosCorrectos[paramName] = valorActual;

                if (isDebugElement)
                {
                    _debugLog.AppendLine($"      Estado: ✓ CORRECTO");
                }
            }
            else
            {
                // Parámetro necesita corrección
                elementData.DatosCompletos = false;
                elementData.ParametrosActualizar[paramName] = valorCorrecto;

                if (isDebugElement)
                {
                    _debugLog.AppendLine($"      Estado: ⚠ NECESITA CORRECCIÓN");
                }
            }
        }

        /// <summary>
        /// Guarda el archivo de debug en el escritorio
        /// </summary>
        private void SaveDebugLog()
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fileName = $"Uniclass_Debug_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                string filePath = Path.Combine(desktopPath, fileName);

                _debugLog.AppendLine();
                _debugLog.AppendLine("=".PadRight(80, '='));
                _debugLog.AppendLine($"Archivo guardado en: {filePath}");
                _debugLog.AppendLine("=".PadRight(80, '='));

                File.WriteAllText(filePath, _debugLog.ToString());

                System.Diagnostics.Debug.WriteLine($"✓ DEBUG LOG guardado en: {filePath}");

                // Mostrar mensaje al usuario
                Autodesk.Revit.UI.TaskDialog.Show("Debug Uniclass",
                    $"Archivo de debug creado:\n\n{filePath}\n\nBúscalo en tu escritorio.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error al guardar debug log: {ex.Message}");
                Autodesk.Revit.UI.TaskDialog.Show("Error Debug",
                    $"No se pudo crear el archivo de debug:\n\n{ex.Message}");
            }
        }
    }
}
