// Models/ElementData.cs
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace TL60_AuditoriaUnificada.Models
{
    /// <summary>
    /// Contiene los datos de un elemento procesado y todos sus resultados de auditoría.
    /// Soporta tanto ViewSchedule (plugin Tablas) como elementos de Revit (plugin COBie).
    /// </summary>
    public class ElementData
    {
        public ElementId ElementId { get; set; }
        public Element Element { get; set; }
        public string Nombre { get; set; }
        public string Categoria { get; set; }
        public string CodigoIdentificacion { get; set; }

        /// <summary>
        /// Se establece en True solo si TODAS las auditorías son correctas.
        /// </summary>
        public bool DatosCompletos { get; set; }

        /// <summary>
        /// Almacena la lista de todos los resultados de la auditoría
        /// (ej. "FILTRO", "COLUMNAS", "RENOMBRAR")
        /// Usado por: Plugin Tablas
        /// </summary>
        public List<AuditItem> AuditResults { get; set; }

        // ============================================
        // PROPIEDADES ESPECÍFICAS DE PLUGIN COBie
        // ============================================

        /// <summary>
        /// Grupo COBie al que pertenece (FACILITY, FLOOR, SPACE, TYPE, COMPONENT)
        /// Usado por: Plugin COBie
        /// </summary>
        public string GrupoCOBie { get; set; }

        /// <summary>
        /// Código de ensamblaje (Assembly Code)
        /// Usado por: Plugin COBie
        /// </summary>
        public string AssemblyCode { get; set; }

        /// <summary>
        /// Indica si todos los parámetros COBie están completos
        /// Usado por: Plugin COBie
        /// </summary>
        public bool CobieCompleto { get; set; }

        /// <summary>
        /// Parámetros que están vacíos
        /// Usado por: Plugin COBie
        /// </summary>
        public List<string> ParametrosVacios { get; set; }

        /// <summary>
        /// Parámetros que están correctos (nombre -> valor)
        /// Usado por: Plugin COBie
        /// </summary>
        public Dictionary<string, string> ParametrosCorrectos { get; set; }

        /// <summary>
        /// Parámetros que necesitan actualizarse (nombre -> valor correcto)
        /// Usado por: Plugin COBie
        /// </summary>
        public Dictionary<string, string> ParametrosActualizar { get; set; }

        /// <summary>
        /// Mensajes de diagnóstico
        /// Usado por: Plugin COBie
        /// </summary>
        public List<string> Mensajes { get; set; }

        public ElementData()
        {
            // Plugin Tablas
            AuditResults = new List<AuditItem>();
            DatosCompletos = false;

            // Plugin COBie
            ParametrosVacios = new List<string>();
            ParametrosCorrectos = new Dictionary<string, string>();
            ParametrosActualizar = new Dictionary<string, string>();
            Mensajes = new List<string>();
            CobieCompleto = false;
        }
    }
}