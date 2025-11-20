// Core/AssemblyCodeExtractor.cs
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TL60_AuditoriaUnificada.Core
{
    /// <summary>
    /// Clase compartida para extraer Assembly Code de elementos.
    /// Utilizada por múltiples plugins (COBie, Uniclass, Metrados, etc.)
    /// </summary>
    public static class AssemblyCodeExtractor
    {
        /// <summary>
        /// Obtiene el Assembly Code de un elemento.
        /// Lógica:
        /// 1. Busca "Assembly Code" en el tipo
        /// 2. Si contiene "TAKEOFF", busca en materiales (MATERIAL_ASSEMBLY CODE)
        /// 3. Si empieza con "C.", usa ese valor
        /// </summary>
        public static string GetAssemblyCode(Element element, Document doc)
        {
            if (element == null) return string.Empty;

            ElementType type = doc.GetElement(element.GetTypeId()) as ElementType;
            if (type == null) return string.Empty;

            Parameter acParam = type.LookupParameter("Assembly Code");
            string assemblyCode = acParam?.AsString()?.Trim() ?? string.Empty;

            // Si contiene "TAKEOFF", buscar en materiales
            if (!string.IsNullOrEmpty(assemblyCode) &&
                assemblyCode.ToUpperInvariant().Contains("TAKEOFF"))
            {
                var materialCodes = GetMaterialAssemblyCodes(element, doc);
                // Retornar el primer código válido encontrado
                return materialCodes.FirstOrDefault() ?? string.Empty;
            }

            // Si empieza con "C.", usar ese valor
            if (!string.IsNullOrEmpty(assemblyCode) &&
                assemblyCode.StartsWith("C.", StringComparison.OrdinalIgnoreCase))
            {
                return assemblyCode;
            }

            return string.Empty;
        }

        /// <summary>
        /// Obtiene todos los Assembly Codes de los materiales de un elemento.
        /// </summary>
        public static IEnumerable<string> GetMaterialAssemblyCodes(Element elem, Document doc)
        {
            var codes = new List<string>();

            if (elem == null || doc == null) return codes;

            // Obtener IDs de materiales del elemento
            ICollection<ElementId> materialIds = elem.GetMaterialIds(false);

            if (materialIds == null || materialIds.Count == 0)
                return codes;

            foreach (ElementId matId in materialIds)
            {
                Material material = doc.GetElement(matId) as Material;
                if (material == null) continue;

                Parameter matAcParam = material.LookupParameter("MATERIAL_ASSEMBLY CODE");
                string matCode = matAcParam?.AsString();

                if (!string.IsNullOrWhiteSpace(matCode) &&
                    matCode.StartsWith("C.", StringComparison.OrdinalIgnoreCase))
                {
                    codes.Add(matCode.Trim());
                }
            }

            return codes;
        }

        /// <summary>
        /// Verifica si un Assembly Code es válido (empieza con "C.")
        /// </summary>
        public static bool IsValidAssemblyCode(string assemblyCode)
        {
            return !string.IsNullOrWhiteSpace(assemblyCode) &&
                   assemblyCode.StartsWith("C.", StringComparison.OrdinalIgnoreCase);
        }
    }
}
