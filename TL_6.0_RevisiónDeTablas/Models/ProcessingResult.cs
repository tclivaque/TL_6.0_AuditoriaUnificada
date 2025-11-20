// Models/ProcessingResult.cs
using System.Collections.Generic;
using System.Linq;

namespace TL60_AuditoriaUnificada.Models
{
    public class ProcessingResult
    {
        public List<ElementData> ElementosProcesados { get; set; }
        public List<string> Errores { get; set; }
        public bool Exitoso { get; set; }
        public string Mensaje { get; set; }

        /// <summary>
        /// Apunta a 'DatosCompletos', que es True solo si TODAS las auditorías son correctas.
        /// </summary>
        public int CantidadCorrectos => ElementosProcesados.Count(e => e.DatosCompletos);

        /// <summary>
        /// Apunta a 'DatosCompletos'.
        /// </summary>
        public int CantidadACorregir => ElementosProcesados.Count(e => !e.DatosCompletos);

        public int CantidadConVacios => 0; // Esta lógica ya no aplica

        public ProcessingResult()
        {
            ElementosProcesados = new List<ElementData>();
            Errores = new List<string>();
            Exitoso = false;
            Mensaje = string.Empty;
        }

        public string GenerarResumen()
        {
            var resumen = $"RESUMEN:\n\n";
            resumen += $"✅ Correctos: {CantidadCorrectos}\n";
            resumen += $"🔧 Tablas con Errores: {CantidadACorregir}\n";
            if (Errores.Count > 0)
            {
                resumen += $"\n❌ Errores: {Errores.Count}\n";
            }
            return resumen;
        }
    }
}