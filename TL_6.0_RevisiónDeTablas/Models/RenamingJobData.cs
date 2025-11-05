// Models/RenamingJobData.cs
using System.Collections.Generic;

namespace TL60_RevisionDeTablas.Models
{
    /// <summary>
    /// Almacena los datos necesarios para la corrección
    /// de una "Tabla de Soporte".
    /// </summary>
    public class RenamingJobData
    {
        public string NuevoNombre { get; set; }
        public string NuevoGrupoVista { get; set; }
        public string NuevoSubGrupoVista { get; set; }
    }
}