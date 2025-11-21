using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using TL60_AuditoriaUnificada.Models;
using TL60_AuditoriaUnificada.Plugins.Tablas;
using TL60_AuditoriaUnificada.Services;

namespace TL60_AuditoriaUnificada.Plugins.Tablas.UI
{
    public partial class TablasPluginControl : UserControl
    {
        private readonly List<DiagnosticRow> _originalDiagnosticRows;
        private List<DiagnosticRow> _diagnosticRows;
        private readonly List<ElementData> _elementosData;
        private readonly Document _doc;
        private readonly ScheduleUpdateAsync _writerAsync;
        private readonly ViewActivatorAsync _viewActivator;

        public TablasPluginControl(
            List<DiagnosticRow> diagnosticRows,
            List<ElementData> elementosData,
            Document doc,
            ScheduleUpdateAsync writerAsync,
            ViewActivatorAsync viewActivator)
        {
            InitializeComponent();
            _originalDiagnosticRows = diagnosticRows ?? new List<DiagnosticRow>();
            _diagnosticRows = _originalDiagnosticRows;
            _elementosData = elementosData;
            _doc = doc;
            _writerAsync = writerAsync;
            _viewActivator = viewActivator;

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
            ErrorTextBlock.Text = $"❌ Errores: {errores}";
            CorrectoTextBlock.Text = $"✅ Correctos: {correctos}";

            DiagnosticDataGrid.ItemsSource = null;
            DiagnosticDataGrid.ItemsSource = _diagnosticRows;
        }

        // ==========================================================
        // ===== MÉTODO CORREGIDO (Switch rellenado)
        // ==========================================================
        private void FilterButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is string filterTag)) return;

            switch (filterTag)
            {
                case "Correcto":
                    _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Correcto).ToList();
                    break;
                case "Advertencia":
                    _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Advertencia).ToList();
                    break;
                case "Corregir":
                    _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Corregir).ToList();
                    break;
                case "Error":
                    _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Error).ToList();
                    break;
                case "Total":
                default:
                    _diagnosticRows = _originalDiagnosticRows;
                    break;
            }

            DiagnosticDataGrid.ItemsSource = null;
            DiagnosticDataGrid.ItemsSource = _diagnosticRows;
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
                    filasSeleccionadas.Any(r => r.ElementId == ed.ElementId)).ToList();

                if (!elementosACorregir.Any())
                    return;

                CorregirButton.IsEnabled = false;
                ExportarButton.IsEnabled = false;

                ProcessingResult writeResult = await Task.Run(() =>
                {
                    return _writerAsync.UpdateSchedulesAsync(_doc, elementosACorregir);
                });

                // Solo mostrar errores
                if (!writeResult.Exitoso)
                {
                    TaskDialog errorDialog = new TaskDialog("Error en Corrección")
                    {
                        MainInstruction = "No se pudieron corregir todos los elementos",
                        MainContent = writeResult.Mensaje ?? "Error desconocido.",
                        ExpandedContent = writeResult.Errores.Any() ?
                            string.Join("\n", writeResult.Errores.Take(10)) : string.Empty
                    };
                    errorDialog.Show();
                }

                CorregirButton.IsEnabled = true;
                ExportarButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Error durante la corrección:\n{ex.Message}");
                CorregirButton.IsEnabled = true;
                ExportarButton.IsEnabled = true;
            }
        }

        private void ExportarButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                FileName = $"Diagnostico_Tablas_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
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

        private void AbrirViewButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var row = button?.DataContext as DiagnosticRow;
            if (row == null || row.ElementId == null || row.ElementId == ElementId.InvalidElementId)
            {
                return;
            }
            row.EsSeleccionable = (row.ElementId != null && row.ElementId != ElementId.InvalidElementId);

            try
            {
                _viewActivator.ActivateViewAsync(row.ElementId);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al intentar abrir la vista: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

    }
}