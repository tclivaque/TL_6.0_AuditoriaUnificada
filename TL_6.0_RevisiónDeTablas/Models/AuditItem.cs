// Models/AuditItem.cs
using TL60_AuditoriaUnificada.Models;

namespace TL60_AuditoriaUnificada.Models
{
    /// <summary>
    /// Almacena el resultado de una única comprobación de auditoría (ej. Filtro, Contenido, etc.)
    /// </summary>
    public class AuditItem
    {
        /// <summary>
        /// El tipo de auditoría (ej: "FILTRO", "CONTENIDO", "CANTIDAD DE COLUMNAS")
        /// </summary>
        public string AuditType { get; set; }

        /// <summary>
        /// Descripción del valor actual encontrado en Revit.
        /// </summary>
        public string ValorActual { get; set; }

        /// <summary>
        /// (¡¡AQUÍ ESTÁ LA CORRECCIÓN!!)
        /// Descripción del valor que se considera correcto.
        /// </summary>
        public string ValorCorregido { get; set; } // <-- ¡¡ESTA LÍNEA ARREGLA LOS 4 ERRORES!!

        /// <summary>
        /// El estado de esta auditoría.
        /// </summary>
        public EstadoParametro Estado { get; set; }

        /// <summary>
        /// Mensaje de error o advertencia.
        /// </summary>
        public string Mensaje { get; set; }

        // --- INICIO DE CORRECCIÓN DE ERRORES ---

        /// <summary>
        /// (NUEVO) Indica si este ítem de auditoría se puede corregir automáticamente.
        /// </summary>
        public bool IsCorrectable { get; set; }

        /// <summary>
        /// (NUEVO) Propiedad genérica para almacenar datos extra (ej. la lista de filtros correctos).
        /// </summary>
        public object Tag { get; set; }

        // --- FIN DE CORRECCIÓN DE ERRORES ---
    }
}