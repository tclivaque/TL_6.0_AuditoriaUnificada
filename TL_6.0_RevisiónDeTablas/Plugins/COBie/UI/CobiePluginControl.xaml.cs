using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TL60_AuditoriaUnificada.Models;
using TL60_AuditoriaUnificada.Plugins.COBie.Services;
using TL60_AuditoriaUnificada.Services;

namespace TL60_AuditoriaUnificada.Plugins.COBie.UI
{
    public partial class CobiePluginControl : UserControl
    {
        private List<DiagnosticRow> _diagnosticRows;
        private List<DiagnosticRow> _originalDiagnosticRows;
        private List<ElementData> _elementosData;
        private Document _doc;
        private ParameterWriterAsync _writerAsync;

        // Procesadores y categorias (se mantienen)
        private FacilityProcessor _facilityProcessor;
        private FloorProcessor _floorProcessor;
        private RoomProcessor _roomProcessor;
        private ElementProcessor _elementProcessor;
        private List<string> _categorias;
        // Campos _config y _definitions eliminados

        public CobiePluginControl(
            List<DiagnosticRow> diagnosticRows,
            List<ElementData> elementosData,
            Document doc,
            FacilityProcessor facilityProcessor = null, FloorProcessor floorProcessor = null,
            RoomProcessor roomProcessor = null, ElementProcessor elementProcessor = null,
            List<string> categorias = null)
        {
            InitializeComponent();

            _originalDiagnosticRows = diagnosticRows ?? new List<DiagnosticRow>();
            _diagnosticRows = _originalDiagnosticRows;
            _elementosData = elementosData;
            _doc = doc;
            _writerAsync = new ParameterWriterAsync();

            _facilityProcessor = facilityProcessor;
            _floorProcessor = floorProcessor;
            _roomProcessor = roomProcessor;
            _elementProcessor = elementProcessor;
            _categorias = categorias;

            LoadData();
        }

        private void LoadData()
        {
            int total = _originalDiagnosticRows.Count;
            int correctos = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Correcto);
            int advertencias = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Advertencia);
            int aCorregir = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Corregir);
            int errores = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Error);

            TotalTextBlock.Text = $"Total: {total}";
            CorregirTextBlock.Text = $"🔧 A Corregir: {aCorregir}";
            AdvertenciaTextBlock.Text = $"⚠ Advertencias: {advertencias}";
            ErrorTextBlock.Text = $"❌ Error: {errores}";
            CorrectoTextBlock.Text = $"✓ Correcto: {correctos}";

            DiagnosticDataGrid.ItemsSource = null;
            DiagnosticDataGrid.ItemsSource = _diagnosticRows;
        }

        private void MostrarButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(sender is Button button) || !(button.Tag is DiagnosticRow row) || !row.EsSeleccionable) return;
                if (row.ElementId != null && row.ElementId != ElementId.InvalidElementId)
                {
                    UIDocument uidoc = new UIDocument(_doc);
                    uidoc.Selection.SetElementIds(new List<ElementId> { row.ElementId });
                    try { uidoc.ShowElements(row.ElementId); } catch { }
                }
            }
            catch { }
        }

        private void HeaderCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var row in _originalDiagnosticRows)
            {
                row.IsChecked = true;
            }
        }

        private void HeaderCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (var row in _originalDiagnosticRows)
            {
                row.IsChecked = false;
            }
        }

        private async void CorregirButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Filtrar solo elementos marcados
                var filasSeleccionadas = _originalDiagnosticRows.Where(r => r.IsChecked && r.Estado == EstadoParametro.Corregir).ToList();
                var elementosACorregir = _elementosData.Where(ed =>
                    filasSeleccionadas.Any(r => r.ElementId == ed.ElementId) && ed.ParametrosActualizar.Any()).ToList();

                if (!elementosACorregir.Any())
                    return;

                CorregirButton.IsEnabled = false;
                CorregirButton.Content = "Corrigiendo...";
                var writerAsync = _writerAsync;
                var doc = _doc;

                ProcessingResult writeResult = await Task.Run(() => writerAsync.WriteParametersAsync(doc, elementosACorregir));

                if (!writeResult.Exitoso)
                {
                    TaskDialog errorDialog = new TaskDialog("Error al Corregir")
                    {
                        MainInstruction = "No se pudieron corregir todos los parámetros",
                        MainContent = writeResult.Mensaje ?? "Error desconocido.",
                        ExpandedContent = writeResult.Errores.Any() ? string.Join("\n", writeResult.Errores) : string.Empty
                    };
                    errorDialog.Show();
                }

                CorregirButton.IsEnabled = true;
                CorregirButton.Content = "Corregir";
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error Inesperado", $"Ocurrió un error:\n{ex.Message}\n{ex.StackTrace}");
                CorregirButton.IsEnabled = true;
                CorregirButton.Content = "Corregir";
            }
        }

        private void FilterButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is string filterTag)) return;
            switch (filterTag)
            {
                case "Correcto": _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Correcto).ToList(); break;
                case "Advertencia": _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Advertencia).ToList(); break;
                case "Corregir": _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Corregir).ToList(); break;
                case "Error": _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Error).ToList(); break;
                case "Total": default: _diagnosticRows = _originalDiagnosticRows; break;
            }
            DiagnosticDataGrid.ItemsSource = null;
            DiagnosticDataGrid.ItemsSource = _diagnosticRows;
        }

        private void ExportarButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                FileName = $"Diagnostico_COBie_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var exportService = new ExcelExportService();
                    byte[] fileBytes = exportService.ExportToExcel(_diagnosticRows);
                    File.WriteAllBytes(dialog.FileName, fileBytes);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error de Exportación", $"Error al exportar el archivo: {ex.Message}");
                }
            }
        }
    }
}