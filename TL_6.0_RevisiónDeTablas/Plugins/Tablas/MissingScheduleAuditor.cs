// Plugins/Tablas/MissingScheduleAuditor.cs
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using TL60_RevisionDeTablas.Core;
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Plugins.Tablas
{
    /// <summary>
    /// Audita los elementos modelados para encontrar Assembly Codes
    /// que deberían tener tablas ("REVIT") pero no la tienen.
    /// </summary>
    public class MissingScheduleAuditor
    {
        private readonly Document _doc;
        private readonly UniclassDataService _uniclassService;

        public MissingScheduleAuditor(Document doc, UniclassDataService uniclassService)
        {
            _doc = doc;
            _uniclassService = uniclassService;
        }

        /// <summary>
        /// Busca tablas faltantes y devuelve un ElementData "dummy" con los errores.
        /// </summary>
        public ElementData FindMissingSchedules(
            IEnumerable<string> existingScheduleCodes,
            List<BuiltInCategory> categoriesToAudit)
        {
            var codesFoundInModel = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            var codesThatNeedSchedule = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);

            var existingCodesSet = new HashSet<string>(existingScheduleCodes, System.StringComparer.OrdinalIgnoreCase);

            // 1. Construir el filtro de categorías
            // Si la lista está vacía, no se auditará nada.
            if (categoriesToAudit == null || categoriesToAudit.Count == 0)
            {
                return null; // No hay nada que auditar
            }
            var categoryFilter = new ElementMulticategoryFilter(categoriesToAudit);

            // 2. Recolectar elementos y sus Assembly Codes
            var collector = new FilteredElementCollector(_doc);
            var elements = collector.WherePasses(categoryFilter).WhereElementIsNotElementType();

            foreach (Element elem in elements)
            {
                ElementType type = _doc.GetElement(elem.GetTypeId()) as ElementType;
                if (type == null) continue;

                // Lógica de Assembly Code (Tipo)
                Parameter acParam = type.LookupParameter("Assembly Code");
                string assemblyCode = acParam?.AsString();

                if (!string.IsNullOrWhiteSpace(assemblyCode) &&
                    assemblyCode.StartsWith("C.", System.StringComparison.OrdinalIgnoreCase))
                {
                    // CASO A: El AC del tipo es "TAKEOFF" -> Buscar en Material
                    if (assemblyCode.ToUpper().Contains("TAKEOFF"))
                    {
                        var materialCodes = GetMaterialAssemblyCodes(elem);
                        foreach (var matCode in materialCodes)
                        {
                            codesFoundInModel.Add(matCode);
                        }
                    }
                    // CASO B: El AC del tipo es normal
                    else
                    {
                        codesFoundInModel.Add(assemblyCode);
                    }
                }
            }

            // 3. Comparar códigos del modelo con la base de datos de Uniclass
            foreach (string code in codesFoundInModel)
            {
                string metradoType = _uniclassService.GetScheduleType(code);

                // Si es "REVIT" y NO existe una tabla para él...
                if (metradoType == "REVIT" && !existingCodesSet.Contains(code))
                {
                    codesThatNeedSchedule.Add(code);
                }
            }

            // 4. Generar el reporte de auditoría
            if (codesThatNeedSchedule.Count > 0)
            {
                // Crear un "ElementData" virtual para reportar errores
                var report = new ElementData
                {
                    ElementId = ElementId.InvalidElementId, // No representa un elemento real
                    Nombre = "Auditoría de Elementos Modelados",
                    Categoria = "Sistema",
                    CodigoIdentificacion = "N/A",
                    DatosCompletos = false
                };

                foreach (string missingCode in codesThatNeedSchedule.OrderBy(c => c))
                {
                    report.AuditResults.Add(new AuditItem
                    {
                        AuditType = "TABLA FALTANTE",
                        Estado = EstadoParametro.Error,
                        Mensaje = $"El Assembly Code '{missingCode}' está asignado a elementos, " +
                                  "está marcado como 'REVIT' en Google Sheets, " +
                                  "pero no se encontró ninguna tabla de planificación para él.",
                        ValorActual = "No existe",
                        ValorCorregido = $"Crear tabla para {missingCode}",
                        IsCorrectable = false
                    });
                }
                return report;
            }

            return null; // No se encontraron errores
        }

        /// <summary>
        /// Obtiene los "MATERIAL_ASSEMBLY CODE" de un elemento.
        /// </summary>
        private IEnumerable<string> GetMaterialAssemblyCodes(Element elem)
        {
            var codes = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            var materialIds = elem.GetMaterialIds(false);

            foreach (ElementId matId in materialIds)
            {
                Material material = _doc.GetElement(matId) as Material;
                if (material == null) continue;

                Parameter matAcParam = material.LookupParameter("MATERIAL_ASSEMBLY CODE");
                string matCode = matAcParam?.AsString();

                if (!string.IsNullOrWhiteSpace(matCode) &&
                    matCode.StartsWith("C.", System.StringComparison.OrdinalIgnoreCase))
                {
                    codes.Add(matCode);
                }
            }
            return codes;
        }
    }
}