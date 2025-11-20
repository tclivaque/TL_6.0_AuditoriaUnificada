using System;

namespace TL60_AuditoriaUnificada.Plugins.COBie.Models
{
    /// <summary>
    /// Representa una definición de parámetro COBie desde Google Sheets "PARÁMETROS COBIE"
    /// </summary>
    public class ParameterDefinition
    {
        /// <summary>
        /// Grupo del parámetro (FACILITY, FLOOR, SPACE, TYPE, COMPONENT, etc.)
        /// </summary>
        public string Grupo { get; set; }

        /// <summary>
        /// Nombre exacto del parámetro en Revit (ej: "COBie.Type.Name")
        /// </summary>
        public string NombreParametro { get; set; }

        /// <summary>
        /// Fecha técnica (metadata)
        /// </summary>
        public string FechaTecnica { get; set; }

        /// <summary>
        /// Valor fijo/por defecto a asignar. Si está vacío, es un valor dinámico.
        /// </summary>
        public string ValorFijo { get; set; }

        /// <summary>
        /// Descripción del parámetro según guía
        /// </summary>
        public string Descripcion { get; set; }

        /// <summary>
        /// Estado del parámetro (LLENO, PENDIENTE, etc.)
        /// </summary>
        public string Estado { get; set; }

        /// <summary>
        /// Indica si el parámetro tiene un valor fijo definido
        /// </summary>
        public bool TieneValorFijo
        {
            get { return !string.IsNullOrWhiteSpace(ValorFijo); }
        }

        /// <summary>
        /// Constructor por defecto
        /// </summary>
        public ParameterDefinition()
        {
            Grupo = string.Empty;
            NombreParametro = string.Empty;
            FechaTecnica = string.Empty;
            ValorFijo = string.Empty;
            Descripcion = string.Empty;
            Estado = string.Empty;
        }

        public override string ToString()
        {
            return $"{Grupo}.{NombreParametro} = {(TieneValorFijo ? ValorFijo : "(dinámico)")}";
        }
    }
}