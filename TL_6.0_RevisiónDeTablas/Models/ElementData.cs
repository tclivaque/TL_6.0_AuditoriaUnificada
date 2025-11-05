// Models/ElementData.cs
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace TL60_RevisionDeTablas.Models
{
    /// <summary>
    /// Contiene los datos de un elemento (ViewSchedule) procesado
    /// y todos sus resultados de auditoría.
    /// </summary>
    public class ElementData
    {
        public ElementId ElementId { get; set; }
        public Element Element { get; set; }
        public string Nombre { get; set; }
        public string Categoria { get; set; }
        public string CodigoIdentificacion { get; set; }

        /// <summary>
        /// Se establece en True solo si TODAS las auditorías (Filtro, Contenido, etc.) son correctas.
        /// </summary>
        public bool DatosCompletos { get; set; }

        /// <summary>
        /// Almacena la lista de todos los resultados de la auditoría 
        /// (ej. "FILTRO", "COLUMNAS", "RENOMBRAR")
        /// </summary>
        public List<AuditItem> AuditResults { get; set; }

        // --- PROPIEDADES ANTIGUAS ELIMINADAS ---
        // 'FiltrosCorrectos' ahora está en AuditItem.Tag
        // 'HeadingsCorregidos' ahora está en AuditItem.Tag
        // 'IsItemizedCorrect' ahora está en AuditItem.Tag


        public ElementData()
        {
            AuditResults = new List<AuditItem>();
            DatosCompletos = false;
        }
    }
}