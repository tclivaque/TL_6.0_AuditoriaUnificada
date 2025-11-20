// Models/RenamingJobData.cs
using System.Collections.Generic;

namespace TL60_AuditoriaUnificada.Models
{
    /// <summary>
    /// Almacena los datos necesarios para la corrección
    /// de una "Tabla de Soporte" o Reclasificación.
    /// </summary>
    public class RenamingJobData
    {
        public string NuevoNombre { get; set; }
        public string NuevoGrupoVista { get; set; }
        public string NuevoSubGrupoVista { get; set; }

        // (¡NUEVO!) Añadido para soportar la lógica de TL_5.0/Diagrama
        public string NuevoSubGrupoVistaSubpartida { get; set; }
    }
}