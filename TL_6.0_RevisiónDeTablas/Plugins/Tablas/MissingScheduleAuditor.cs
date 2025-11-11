// Plugins/Tablas/MissingScheduleAuditor.cs
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System;
using TL60_RevisionDeTablas.Core;
using TL60_RevisionDeTablas.Models;
using System.Text;
using System.IO;

namespace TL60_RevisionDeTablas.Plugins.Tablas
{
    public class MissingScheduleAuditor
    {
        private readonly Document _doc;
        private readonly UniclassDataService _uniclassService;

        public MissingScheduleAuditor(Document doc, UniclassDataService uniclassService)
        {
            _doc = doc;
            _uniclassService = uniclassService;
        }

        public ElementData FindMissingSchedules(
            IEnumerable<string> existingScheduleCodes,
            List<BuiltInCategory> categoriesToAudit,
            StringBuilder sb)
        {
            sb.AppendLine("  --- Debugging MissingScheduleAuditor ---");

            // 1. Obtener Especialidad del Documento Principal
            string mainDocTitle = Path.GetFileNameWithoutExtension(_doc.Title);
            string mainSpecialty = "UNKNOWN_SPECIALTY";
            try
            {
                var parts = mainDocTitle.Split('-');
                if (parts.Length > 3)
                {
                    mainSpecialty = parts[3];
                }
                sb.AppendLine($"  > Especialidad del Doc. Principal: '{mainSpecialty}' (de '{mainDocTitle}')");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"  > ADVERTENCIA: No se pudo extraer especialidad de '{mainDocTitle}': {ex.Message}");
            }

            // 2. Obtener lista de TODOS los documentos a escanear (Principal + Vínculos)
            List<Document> docsToScan = new List<Document>();
            docsToScan.Add(_doc);
            sb.AppendLine("  > Documento principal añadido a la cola de escaneo.");

            sb.AppendLine("  > Buscando Vínculos (Revit Links) relevantes...");
            var linkInstances = new FilteredElementCollector(_doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>();

            int linksEncontrados = 0;
            int linksIgnorados = 0;
            int linksRelevantes = 0;

            foreach (var linkInstance in linkInstances)
            {
                Document linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc == null) continue;

                linksEncontrados++;
                RevitLinkType linkType = _doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType;
                if (linkType == null) continue;

                string linkName = linkType.Name;
                string linkSpecialty = "UNKNOWN_LINK_SPECIALTY";

                try
                {
                    var parts = Path.GetFileNameWithoutExtension(linkName).Split('-');
                    if (parts.Length > 3)
                    {
                        linkSpecialty = parts[3];
                    }
                }
                catch
                {
                    sb.AppendLine($"    > Vínculo '{linkName}' tiene nombre no estándar.");
                    continue;
                }

                if (linkSpecialty.Equals(mainSpecialty, StringComparison.OrdinalIgnoreCase))
                {
                    if (!docsToScan.Contains(linkDoc))
                    {
                        docsToScan.Add(linkDoc);
                        linksRelevantes++;
                    }
                }
                else
                {
                    linksIgnorados++;
                }
            }
            sb.AppendLine($"  > Vínculos encontrados: {linksEncontrados} (Cargados)");
            sb.AppendLine($"  > Vínculos RELEVANTES (misma especialidad '{mainSpecialty}'): {linksRelevantes}");
            sb.AppendLine($"  > Vínculos Ignorados (otra especialidad): {linksIgnorados}");
            sb.AppendLine($"  > Total de documentos a escanear (Principal + Vínculos): {docsToScan.Count}");

            // ==========================================================
            // 3. Escanear TODOS los documentos de la lista
            // (¡MODIFICADO! Usar un Dictionary para guardar el origen del AC)
            // ==========================================================
            var codesFoundInModel = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
            var codesThatNeedSchedule = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
            var existingCodesSet = new HashSet<string>(existingScheduleCodes, System.StringComparer.OrdinalIgnoreCase);
            var categoryFilter = new ElementMulticategoryFilter(categoriesToAudit);

            foreach (Document scanDoc in docsToScan)
            {
                string docName = Path.GetFileName(scanDoc.PathName);
                if (string.IsNullOrEmpty(docName)) docName = "Documento Principal";

                sb.AppendLine($"  --- Escaneando Doc: {docName} ---");

                var elementsCollector = new FilteredElementCollector(scanDoc)
                                        .WherePasses(categoryFilter)
                                        .WhereElementIsNotElementType();

                int elementCount = 0;
                int acFoundCount = 0;

                foreach (Element elem in elementsCollector.ToElements())
                {
                    elementCount++;
                    ElementType type = scanDoc.GetElement(elem.GetTypeId()) as ElementType;
                    if (type == null) continue;

                    Parameter acParam = type.LookupParameter("Assembly Code");
                    string assemblyCode = acParam?.AsString();

                    if (!string.IsNullOrWhiteSpace(assemblyCode) &&
                        assemblyCode.StartsWith("C.", System.StringComparison.OrdinalIgnoreCase))
                    {
                        if (assemblyCode.ToUpper().Contains("TAKEOFF"))
                        {
                            var materialCodes = GetMaterialAssemblyCodes(elem, scanDoc);
                            foreach (var matCode in materialCodes)
                            {
                                // (¡MODIFICADO!) Guardar AC y docName
                                if (!codesFoundInModel.ContainsKey(matCode))
                                {
                                    codesFoundInModel.Add(matCode, docName);
                                    acFoundCount++;
                                }
                            }
                        }
                        else
                        {
                            // (¡MODIFICADO!) Guardar AC y docName
                            if (!codesFoundInModel.ContainsKey(assemblyCode))
                            {
                                codesFoundInModel.Add(assemblyCode, docName);
                                acFoundCount++;
                            }
                        }
                    }
                }
                sb.AppendLine($"  > Elementos analizados: {elementCount}");
                sb.AppendLine($"  > Assembly Codes únicos (nuevos) encontrados: {acFoundCount}");
            }
            sb.AppendLine($"  --- Fin del Escaneo ---");

            // ==========================================================
            // 4. Comparar y Reportar (¡MODIFICADO! Propagar el docName)
            // ==========================================================
            sb.AppendLine("  > Comparando AC encontrados con Google Sheets y Tablas Existentes...");
            int missingCount = 0;
            foreach (var kvp in codesFoundInModel) // kvp.Key = AC, kvp.Value = docName
            {
                string code = kvp.Key;
                string sourceDoc = kvp.Value;
                string metradoType = _uniclassService.GetScheduleType(code);

                if (metradoType == "REVIT" && !existingCodesSet.Contains(code))
                {
                    codesThatNeedSchedule.Add(code, sourceDoc); // Guardar AC y docName
                    missingCount++;
                }
            }
            sb.AppendLine($"  > Total de AC 'REVIT' sin tabla: {missingCount}");

            // ==========================================================
            // 5. Generar el reporte de auditoría (¡MODIFICADO!)
            // ==========================================================
            if (codesThatNeedSchedule.Count > 0)
            {
                var report = new ElementData
                {
                    ElementId = ElementId.InvalidElementId,
                    Nombre = "Auditoría de Elementos (Principal + Vínculos)",
                    Categoria = "Sistema",
                    CodigoIdentificacion = "N/A",
                    DatosCompletos = false
                };

                foreach (var missingItem in codesThatNeedSchedule.OrderBy(kvp => kvp.Key))
                {
                    string missingCode = missingItem.Key;
                    string rvtName = missingItem.Value;

                    report.AuditResults.Add(new AuditItem
                    {
                        AuditType = "TABLA FALTANTE",
                        // (¡MODIFICADO!) Cambiado de Error a Vacio (Advertencia)
                        Estado = EstadoParametro.Vacio,
                        // (¡MODIFICADO!) Nuevo formato de mensaje
                        Mensaje = $"No se encontró tabla para el AC: '{missingCode}' (detectado en el modelo: '{rvtName}')",
                        ValorActual = "No existe",
                        ValorCorregido = $"Crear tabla para {missingCode}",
                        IsCorrectable = false
                    });
                }
                return report;
            }

            return null; // No se encontraron errores
        }

        private IEnumerable<string> GetMaterialAssemblyCodes(Element elem, Document doc)
        {
            var codes = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            var materialIds = elem.GetMaterialIds(false);

            foreach (ElementId matId in materialIds)
            {
                Material material = doc.GetElement(matId) as Material;
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