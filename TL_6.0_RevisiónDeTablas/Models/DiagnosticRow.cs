// Models/DiagnosticRow.cs
using System.Windows.Media;
using Autodesk.Revit.DB;

namespace TL60_AuditoriaUnificada.Models
{
    /// <summary>
    /// Modelo para una fila del diagnóstico (tabla de la ventana).
    /// Soporta tanto plugin Tablas como plugin COBie.
    /// </summary>
    public class DiagnosticRow
    {
        public int NumeroFila { get; set; }
        public bool EsSeleccionable { get; set; }
        public ElementId ElementId { get; set; }
        public string IdMostrar { get; set; }

        /// <summary>
        /// Grupo para agrupar filas: "UNIDADES GLOBALES" o "TABLAS"
        /// Usado por: Plugin Tablas
        /// </summary>
        public string Grupo { get; set; }

        /// <summary>
        /// Grupo COBie: "FACILITY", "FLOOR", "SPACE", "TYPE", "COMPONENT"
        /// Usado por: Plugin COBie
        /// </summary>
        public string GrupoCOBie { get; set; }

        /// <summary>
        /// Código de ensamblaje (Assembly Code)
        /// Usado por: Plugin COBie
        /// </summary>
        public string AssemblyCode { get; set; }

        public string CodigoIdentificacion { get; set; }
        public string Descripcion { get; set; }
        public string NombreParametro { get; set; }
        public string ValorActual { get; set; }
        public string ValorCorregido { get; set; }
        public EstadoParametro Estado { get; set; }
        public string Mensaje { get; set; }
        public SolidColorBrush BackgroundColor { get; set; }
    }

    /// <summary>
    /// Estados posibles de un parámetro
    /// </summary>
    public enum EstadoParametro
    {
        Correcto,   // Verde
        Vacio,      // Naranja
        Corregir,   // Azul
        Error       // Rojo
    }
}