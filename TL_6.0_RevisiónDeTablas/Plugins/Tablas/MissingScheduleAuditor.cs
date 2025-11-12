// Plugins/Tablas/MissingScheduleAuditor.cs
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System;
using TL60_RevisionDeTablas.Core;
using TL60_RevisionDeTablas.Models;
using System.IO;

namespace TL60_RevisionDeTablas.Plugins.Tablas
{
    public class MissingScheduleAuditor
    {
        private readonly Document _doc;
        private readonly UniclassDataService _uniclassService;

        // ==========================================================
        // ===== (¡NUEVO!) Lista Blanca de Equipamiento (EM)
        // ==========================================================
        private static readonly List<string> EM_WHITELIST = new List<string>
        {
            "200114-CCC02-MO-EM-000410",
            "200114-CCC02-MO-EM-045500",
            "200114-CCC02-MO-EM-045600"
        };
        // ==========================================================

        public MissingScheduleAuditor(Document doc, UniclassDataService uniclassService)
        {
            _doc = doc;
            _uniclassService = uniclassService;
        }

        /// <summary>
        /// Busca tablas faltantes y devuelve un ElementData "dummy" con los errores.
        /// (¡MODIFICADO!) Ahora devuelve los AC encontrados y aplica la lógica de "EM Whitelist".
        /// </summary>
        public ElementData FindMissingSchedules(
            IEnumerable<string> existingScheduleCodes,
            List<BuiltInCategory> categoriesToAudit,
            out HashSet<string> allFoundCodes) // <-- (¡MODIFICADO!) Parámetro de salida
        {
            var codesFoundInModel = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var codesThatNeedSchedule = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var existingCodesSet = new HashSet<string>(existingScheduleCodes, StringComparer.OrdinalIgnoreCase);

            // 1. Obtener Especialidad del Documento Principal
            string mainDocTitle = Path.GetFileNameWithoutExtension(_doc.Title);
            string mainSpecialty = GetSpecialtyFromTitle(mainDocTitle);

            // 2. Obtener lista de TODOS los documentos a escanear (Principal + Vínculos)
            List<Document> docsToScan = GetDocumentsToScan(mainSpecialty);

            // 3. Escanear TODOS los documentos de la lista
            var categoryFilter = new ElementMulticategoryFilter(categoriesToAudit);

            foreach (Document scanDoc in docsToScan)
            {
                string docName = Path.GetFileName(scanDoc.PathName);
                if (string.IsNullOrEmpty(docName)) docName = "Documento Principal";

                var elementsCollector = new FilteredElementCollector(scanDoc)
                                        .WherePasses(categoryFilter)
                                        .WhereElementIsNotElementType();

                foreach (Element elem in elementsCollector.ToElements())
                {
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
                                if (!codesFoundInModel.ContainsKey(matCode))
                                {
                                    codesFoundInModel.Add(matCode, docName);
                                }
                            }
                        }
                        else
                        {
                            if (!codesFoundInModel.ContainsKey(assemblyCode))
                            {
                                codesFoundInModel.Add(assemblyCode, docName);
                            }
                        }
                    }
                }
            }

            // (¡MODIFICADO!) Devolver todos los códigos encontrados
            allFoundCodes = new HashSet<string>(codesFoundInModel.Keys);

            // 4. Comparar y Reportar (Lógica sin cambios)
            foreach (var kvp in codesFoundInModel) // kvp.Key = AC, kvp.Value = docName
            {
                string code = kvp.Key;
                string sourceDoc = kvp.Value;
                string metradoType = _uniclassService.GetScheduleType(code);

                if (metradoType == "REVIT" && !existingCodesSet.Contains(code))
                {
                    codesThatNeedSchedule.Add(code, sourceDoc); // Guardar AC y docName
                }
            }

            // 5. Generar el reporte de auditoría (Lógica sin cambios)
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
                        Estado = EstadoParametro.Vacio, // Advertencia
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

        /// <summary>
        /// (¡NUEVO!) Implementa la lógica de filtrado de documentos (Modo Normal vs Modo EM)
        /// </summary>
        private List<Document> GetDocumentsToScan(string mainSpecialty)
        {
            List<Document> docsToScan = new List<Document>();

            // ==========================================================
            // ===== CASO A: Modo Restringido EM
            // ==========================================================
            if (mainSpecialty.Equals("EM", StringComparison.OrdinalIgnoreCase))
            {
                // 1. Validar anfitrión
                string mainDocTitle = Path.GetFileNameWithoutExtension(_doc.Title);
                if (EM_WHITELIST.Contains(mainDocTitle))
                {
                    docsToScan.Add(_doc);
                }

                // 2. Validar Vínculos
                var linkInstances = new FilteredElementCollector(_doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();

                foreach (var linkInstance in linkInstances)
                {
                    Document linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc == null) continue;

                    RevitLinkType linkType = _doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType;
                    if (linkType == null) continue;

                    string linkDocTitle = Path.GetFileNameWithoutExtension(linkType.Name);

                    // Solo añadir si el nombre del vínculo está en la lista blanca
                    if (EM_WHITELIST.Contains(linkDocTitle))
                    {
                        if (!docsToScan.Contains(linkDoc))
                        {
                            docsToScan.Add(linkDoc);
                        }
                    }
                }
            }
            // ==========================================================
            // ===== CASO B: Modo Normal (EE, AR, ES, etc.)
            // ==========================================================
            else
            {
                docsToScan.Add(_doc); // Añadir anfitrión

                var linkInstances = new FilteredElementCollector(_doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();

                foreach (var linkInstance in linkInstances)
                {
                    Document linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc == null) continue;

                    RevitLinkType linkType = _doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType;
                    if (linkType == null) continue;

                    string linkName = linkType.Name;
                    string linkSpecialty = GetSpecialtyFromTitle(linkName);

                    // Lógica original: Añadir si la especialidad coincide
                    if (linkSpecialty.Equals(mainSpecialty, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!docsToScan.Contains(linkDoc))
                        {
                            docsToScan.Add(linkDoc);
                        }
                    }
                }
            }

            return docsToScan;
        }

        /// <summary>
        /// Helper para extraer la especialidad (ej. "EE") de un nombre de archivo.
        /// </summary>
        private string GetSpecialtyFromTitle(string docTitle)
        {
            try
            {
                var parts = Path.GetFileNameWithoutExtension(docTitle).Split('-');
                if (parts.Length > 3)
                {
                    return parts[3]; // "ES", "EE", "AR", etc.
                }
            }
            catch (Exception) { /* Ignorar */ }
            return "UNKNOWN_SPECIALTY";
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