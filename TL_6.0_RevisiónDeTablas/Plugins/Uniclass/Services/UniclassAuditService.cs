// Plugins/Uniclass/Services/UniclassAuditService.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using TL60_AuditoriaUnificada.Core;
using TL60_AuditoriaUnificada.Models;

namespace TL60_AuditoriaUnificada.Plugins.Uniclass.Services
{
    /// <summary>
    /// Servicio de auditoría de parámetros Uniclass
    /// </summary>
    public class UniclassAuditService
    {
        private readonly Document _doc;
        private readonly UniclassDataService _uniclassService;
        private readonly List<BuiltInCategory> _categoriasAuditar;

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
        }

        /// <summary>
        /// Procesa todos los elementos del proyecto y genera datos de auditoría
        /// </summary>
        public List<ElementData> ProcessElements()
        {
            var elementosData = new List<ElementData>();

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

            // Agrupar elementos por tipo para procesar por tipo
            var elementosPorTipo = elementos
                .Where(e => e.GetTypeId() != null && e.GetTypeId() != ElementId.InvalidElementId)
                .GroupBy(e => e.GetTypeId());

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

            return elementosData;
        }

        /// <summary>
        /// Procesa un tipo de elemento específico
        /// </summary>
        private ElementData ProcessElementType(ElementType tipo, string assemblyCode, int instanciasCount)
        {
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
            Dictionary<string, string> valoresCorrectos = GetValoresCorrectos(assemblyCode);

            // Auditar cada uno de los 6 parámetros
            foreach (var paramName in UNICLASS_PARAMETERS)
            {
                AuditarParametro(tipo, paramName, assemblyCode, valoresCorrectos, elementData);
            }

            return elementData;
        }

        /// <summary>
        /// Determina los valores correctos para los parámetros Uniclass según el Assembly Code
        /// </summary>
        private Dictionary<string, string> GetValoresCorrectos(string assemblyCode)
        {
            var valoresCorrectos = new Dictionary<string, string>();

            // Si el Assembly Code NO empieza con "C.", llenar con "-"
            if (!AssemblyCodeExtractor.IsValidAssemblyCode(assemblyCode))
            {
                foreach (var paramName in UNICLASS_PARAMETERS)
                {
                    valoresCorrectos[paramName] = "-";
                }
                return valoresCorrectos;
            }

            // Assembly Code válido (empieza con "C.")
            // Buscar en la matriz Uniclass
            var uniclassData = _uniclassService.GetUniclassParameters(assemblyCode);

            if (uniclassData == null)
            {
                // No se encontró en la matriz, marcar error
                foreach (var paramName in UNICLASS_PARAMETERS)
                {
                    valoresCorrectos[paramName] = $"ERROR: {assemblyCode} no encontrado en matriz";
                }
                return valoresCorrectos;
            }

            // Verificar columna UNICLAS
            if (!uniclassData.ShouldAudit)
            {
                // UNICLAS no tiene "✓", llenar con "-"
                foreach (var paramName in UNICLASS_PARAMETERS)
                {
                    valoresCorrectos[paramName] = "-";
                }
                return valoresCorrectos;
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
            ElementData elementData)
        {
            // Obtener valor actual del parámetro
            Parameter param = tipo.LookupParameter(paramName);

            if (param == null)
            {
                // Parámetro no existe en el tipo
                elementData.DatosCompletos = false;
                elementData.Mensajes.Add($"Error: El tipo '{tipo.Name}' no tiene el parámetro '{paramName}'");
                return;
            }

            string valorActual = param.AsString() ?? string.Empty;
            string valorCorrecto = valoresCorrectos.ContainsKey(paramName) ? valoresCorrectos[paramName] : string.Empty;

            // Caso especial: valor correcto es un mensaje de error
            if (valorCorrecto.StartsWith("ERROR:"))
            {
                elementData.DatosCompletos = false;
                elementData.Mensajes.Add(valorCorrecto);
                return;
            }

            // Comparar valores
            if (valorActual == valorCorrecto)
            {
                // Parámetro correcto
                elementData.ParametrosCorrectos[paramName] = valorActual;
            }
            else
            {
                // Parámetro necesita corrección
                elementData.DatosCompletos = false;
                elementData.ParametrosActualizar[paramName] = valorCorrecto;
            }
        }
    }
}
