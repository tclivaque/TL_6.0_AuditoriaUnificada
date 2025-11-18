using System;
using System.Collections.Generic;

namespace TL60_RevisionDeTablas.Plugins.COBie.Models
{
    /// <summary>
    /// Modelo de configuración leída desde Google Sheets "ENTRADAS_SCRIPT_2.0 COBIE"
    /// </summary>
    public class CobieConfig
    {
        /// <summary>
        /// ID del Google Spreadsheet (14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y)
        /// </summary>
        public string SpreadsheetId { get; set; }

        /// <summary>
        /// Nombre de la hoja con datos de Assembly Code (default: "ASSEMBLY CODE COBIE")
        /// </summary>
        public string NombreHojaAssemblyCode { get; set; }

        /// <summary>
        /// Nombre de la hoja con datos de Rooms (default: "AMBIENTES ÚNICOS")
        /// </summary>
        public string NombreHojaRooms { get; set; }

        /// <summary>
        /// Nombre de la hoja con datos de Levels (default: "NIVELES")
        /// </summary>
        public string NombreHojaLevels { get; set; }

        /// <summary>
        /// Rango de columnas a procesar para COBie (ej: "H1:AF1")
        /// </summary>
        public string RangoColumnasCOBie { get; set; }

        /// <summary>
        /// Lista de categorías de Revit a procesar (ej: "OST_PlumbingFixtures")
        /// </summary>
        public List<string> Categorias { get; set; }

        // ===== CAMBIOS AQUÍ: Nuevas propiedades para valores por defecto =====

        /// <summary>
        /// Valor por defecto para parámetros String sin Assembly Code
        /// </summary>
        public string ValorDefectoString { get; set; }

        /// <summary>
        /// Valor por defecto para parámetros Integer Yes/No sin Assembly Code
        /// </summary>
        public string ValorDefectoIntegerYesNo { get; set; }

        /// <summary>
        /// Valor por defecto para parámetros Integer numéricos sin Assembly Code
        /// </summary>
        public string ValorDefectoInteger { get; set; }

        /// <summary>
        /// Valor por defecto para parámetros Double sin Assembly Code
        /// </summary>
        public string ValorDefectoDouble { get; set; }

        /// <summary>
        /// Constructor con valores por defecto
        /// </summary>
        public CobieConfig()
        {
            SpreadsheetId = string.Empty;
            NombreHojaAssemblyCode = "ASSEMBLY CODE COBIE";
            NombreHojaRooms = "AMBIENTES ÚNICOS";
            RangoColumnasCOBie = "H1:AF1";
            Categorias = new List<string>();

            // Valores por defecto (ahora se sobreescribirán desde G-Sheets)
            ValorDefectoString = "n/a";
            ValorDefectoIntegerYesNo = "1";
            ValorDefectoInteger = "0";
            ValorDefectoDouble = "0";
            NombreHojaLevels = "NIVELES";
        }
    }
}