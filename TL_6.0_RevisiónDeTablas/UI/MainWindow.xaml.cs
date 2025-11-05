// UI/MainWindow.xaml.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using TL60_RevisionDeTablas.Models;
using TL60_RevisionDeTablas.Services;
using System.IO;

namespace TL60_RevisionDeTablas.UI
{
    public partial class MainWindow : Window
    {
        // (NUEVO) Almacena la lista original sin filtrar
        private readonly List<DiagnosticRow> _originalDiagnosticRows;
        // (MODIFICADO) Esta lista ahora contendrá los datos filtrados
        private List<DiagnosticRow> _diagnosticRows;

        private readonly List<ElementData> _elementosData;
        private readonly Document _doc;
        private readonly ScheduleUpdateAsync _writerAsync;
        private readonly ViewActivatorAsync _viewActivator;

        public MainWindow(
            List<DiagnosticRow> diagnosticRows,
            List<ElementData> elementosData,
            Document doc,
            ScheduleUpdateAsync writerAsync,
            ViewActivatorAsync viewActivator)
        {
            InitializeComponent();
            _originalDiagnosticRows = diagnosticRows ?? new List<DiagnosticRow>();
            _diagnosticRows = _originalDiagnosticRows; // Al inicio, la lista mostrada es la original
            _elementosData = elementosData;
            _doc = doc;
            _writerAsync = writerAsync;
            _viewActivator = viewActivator;

            LoadData();
        }

        private void LoadData()
        {
            // (MODIFICADO) Lógica de conteo copiada del proyecto COBie
            int total = _originalDiagnosticRows.Count;
            int correctos = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Correcto);
            int vacios = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Vacio);
            int aCorregir = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Corregir);
            int errores = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Error);

            // (MODIFICADO) Actualizar los TextBlocks de la nueva leyenda
            TotalTextBlock.Text = $"Total: {total}";
            CorregirTextBlock.Text = $"🔧 A Corregir: {aCorregir}";
            VacioTextBlock.Text = $"⚠ Advertencias: {vacios}";
            ErrorTextBlock.Text = $"❌ Errores: {errores}";
            CorrectoTextBlock.Text = $"✅ Correctos: {correctos}";

            // (MODIFICADO) El DataGrid ahora se alimenta de la lista _diagnosticRows (que puede ser filtrada)
            DiagnosticDataGrid.ItemsSource = null;
            DiagnosticDataGrid.ItemsSource = _diagnosticRows;
        }

        /// <summary>
        /// (NUEVO) Manejador de clics para los botones de la leyenda (copiado de COBie)
        /// </summary>
        private void FilterButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is string filterTag)) return;

            // Filtrar la lista original y asignarla a la lista mostrada
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

            // Recargar el DataGrid con la lista filtrada
            DiagnosticDataGrid.ItemsSource = null;
            DiagnosticDataGrid.ItemsSource = _diagnosticRows;
        }

        /// <summary>
        /// Botón "Corregir"
        /// </summary>
        private async void CorregirButton_Click(object sender, RoutedEventArgs e)
        {
            // (Lógica sin cambios, ya era correcta)
            var idsACorregir = _originalDiagnosticRows // (MODIFICADO) Siempre buscar en la lista original
                .Where(r => r.Estado == EstadoParametro.Corregir)
                .Select(r => r.ElementId)
                .Distinct()
                .ToList();

            if (idsACorregir.Count == 0)
            {
                MessageBox.Show("No hay elementos con correcciones aplicables (FILTRO o CONTENIDO).", "Información", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var elementosACorregir = _elementosData
                .Where(ed => idsACorregir.Contains(ed.ElementId))
                .ToList();

            try
            {
                CorregirButton.IsEnabled = false;
                CerrarButton.IsEnabled = false; // (MODIFICADO) Deshabilitar ambos botones
                CorregirButton.Content = "Corrigiendo...";

                ProcessingResult writeResult = await Task.Run(() =>
                {
                    return _writerAsync.UpdateSchedulesAsync(_doc, elementosACorregir);
                });

                CorregirButton.IsEnabled = true;
                CerrarButton.IsEnabled = true; // (MODIFICADO) Rehabilitar ambos botones
                CorregirButton.Content = "Corregir";

                if (!writeResult.Exitoso)
                {
                    MessageBox.Show(writeResult.Mensaje, "Error al Corregir", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    MessageBox.Show(writeResult.Mensaje, "Éxito", MessageBoxButton.OK, MessageBoxImage.Information);
                    this.Close(); // Cerrar al éxito
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error fatal: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                CorregirButton.IsEnabled = true;
                CerrarButton.IsEnabled = true; // (MODIFICADO) Rehabilitar ambos botones
                CorregirButton.Content = "Corregir";
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
                    // (MODIFICADO) Exportar la lista que se está mostrando actualmente (filtrada o no)
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
            // (NUEVO) Asignar propiedad EsSeleccionable de la fila, copiado de COBie
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

        private void CerrarButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}