using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using TL60_RevisionDeTablas.Models;
using TL60_RevisionDeTablas.Plugins.Tablas;

namespace TL60_RevisionDeTablas.Plugins.Tablas.UI
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
            int vacios = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Vacio);
            int aCorregir = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Corregir);
            int errores = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Error);

            TotalTextBlock.Text = $"Total: {total}";
            CorregirTextBlock.Text = $"🔧 A Corregir: {aCorregir}";
            VacioTextBlock.Text = $"⚠ Advertencias: {vacios}";
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
                case "Vacio":
                    _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Vacio).ToList();
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

        // --- MÉTODO DE CORRECIÓN (Ya corregido en la respuesta anterior) ---
        private async void CorregirButton_Click(object sender, RoutedEventArgs e)
        {
            CorregirButton.IsEnabled = false;
            ExportarButton.IsEnabled = false;

            try
            {
                ProcessingResult writeResult = await Task.Run(() =>
                {
                    // Llamar al handler con la lista COMPLETA
                    return _writerAsync.UpdateSchedulesAsync(_doc, _elementosData);
                });

                // Mostrar el resultado combinado
                bool huboErrores = !writeResult.Exitoso;
                string mensaje = writeResult.Mensaje;

                if (huboErrores && writeResult.Errores.Count > 0)
                {
                    mensaje += "\n\nErrores:\n" + string.Join("\n", writeResult.Errores.Take(10));
                }

                if (string.IsNullOrWhiteSpace(mensaje))
                {
                    mensaje = "No se encontraron elementos para corregir.";
                }

                MessageBox.Show(mensaje, "Resultado de Corrección", MessageBoxButton.OK,
                    huboErrores ? MessageBoxImage.Error : MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error durante la corrección:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
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

                    MessageBox.Show($"Reporte exportado exitosamente a:\n{dialog.FileName}", "Exportación Completa", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al exportar el archivo: {ex.Message}", "Error de Exportación", MessageBoxButton.OK, MessageBoxImage.Error);
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