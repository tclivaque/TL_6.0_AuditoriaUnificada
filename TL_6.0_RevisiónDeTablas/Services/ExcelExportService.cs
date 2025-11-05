// Services/ExcelExportService.cs
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using TL60_RevisionDeTablas.Models;

namespace TL60_RevisionDeTablas.Services
{
    public class ExcelExportService
    {
        public ExcelExportService()
        {
            // (REVERTIDO) Volver al código original.
            // Esto dará una advertencia, pero compilará.
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public byte[] ExportToExcel(List<DiagnosticRow> diagnosticRows)
        // ... (el resto del archivo queda igual)        // ... (el resto del archivo queda igual)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("DiagnosticoFiltros");

                // Encabezados
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "ASSEMBLY CODE";
                worksheet.Cells[1, 3].Value = "NOMBRE DE TABLA";

                // ===================================================
                // ===== CORRECCIÓN #5: Encabezado de Excel alineado =====
                // ===================================================
                worksheet.Cells[1, 4].Value = "AUDITORÍA"; // (Ya estaba correcto)

                worksheet.Cells[1, 5].Value = "VALOR ACTUAL";
                worksheet.Cells[1, 6].Value = "VALOR CORRECTO";
                worksheet.Cells[1, 7].Value = "ESTADO";
                worksheet.Cells[1, 8].Value = "MENSAJE";

                // Aplicar formato a los encabezados
                using (var range = worksheet.Cells[1, 1, 1, 8])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // Escribir datos
                for (int i = 0; i < diagnosticRows.Count; i++)
                {
                    var row = diagnosticRows[i];
                    int rowIndex = i + 2;

                    worksheet.Cells[rowIndex, 1].Value = row.IdMostrar;
                    worksheet.Cells[rowIndex, 2].Value = row.CodigoIdentificacion;
                    worksheet.Cells[rowIndex, 3].Value = row.Descripcion;
                    worksheet.Cells[rowIndex, 4].Value = row.NombreParametro;
                    worksheet.Cells[rowIndex, 5].Value = row.ValorActual;
                    worksheet.Cells[rowIndex, 6].Value = row.ValorCorregido;
                    worksheet.Cells[rowIndex, 7].Value = row.Estado.ToString();
                    worksheet.Cells[rowIndex, 8].Value = row.Mensaje;
                }

                worksheet.Cells[1, 1, 1, 8].AutoFitColumns();

                return package.GetAsByteArray();
            }
        }
    }
}