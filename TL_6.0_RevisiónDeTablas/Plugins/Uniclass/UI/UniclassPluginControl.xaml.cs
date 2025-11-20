// Plugins/Uniclass/UI/UniclassPluginControl.xaml.cs
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
using TL60_AuditoriaUnificada.Plugins.Uniclass.Services;
using TL60_AuditoriaUnificada.Services;

namespace TL60_AuditoriaUnificada.Plugins.Uniclass.UI
{
    public partial class UniclassPluginControl : UserControl
    {
        private readonly UIDocument _uidoc;
        private readonly Document _doc;
        private List<DiagnosticRow> _todasLasFilas;
        private List<ElementData> _elementosData;
        private string _filtroActual = "Total";
        private readonly UniclassParameterWriterAsync _parameterWriterAsync;

        public UniclassPluginControl(UIDocument uidoc, List<ElementData> elementosData)
        {
            InitializeComponent();
            _uidoc = uidoc ?? throw new ArgumentNullException(nameof(uidoc));
            _doc = uidoc.Document;
            _elementosData = elementosData ?? new List<ElementData>();
            _todasLasFilas = new List<DiagnosticRow>();
            _parameterWriterAsync = new UniclassParameterWriterAsync();

            CargarDatos();
        }

        private void CargarDatos()
        {
            try
            {
                // Construir filas de diagnóstico
                var builder = new UniclassDataBuilder();
                _todasLasFilas = builder.BuildDiagnosticRows(_elementosData);

                // Actualizar contadores
                ActualizarContadores();

                // Mostrar todas las filas inicialmente
                DiagnosticDataGrid.ItemsSource = _todasLasFilas;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error de Carga", $"Error al cargar datos de Uniclass:\n\n{ex.Message}");
            }
        }

        private void ActualizarContadores()
        {
            int total = _todasLasFilas.Count;
            int corregir = _todasLasFilas.Count(r => r.Estado == EstadoParametro.Corregir);
            int advertencia = _todasLasFilas.Count(r => r.Estado == EstadoParametro.Advertencia);
            int error = _todasLasFilas.Count(r => r.Estado == EstadoParametro.Error);
            int correcto = _todasLasFilas.Count(r => r.Estado == EstadoParametro.Correcto);

            TotalTextBlock.Text = $"Total: {total}";
            CorregirTextBlock.Text = $"🔧 A Corregir: {corregir}";
            AdvertenciaTextBlock.Text = $"⚠ Advertencias: {advertencia}";
            ErrorTextBlock.Text = $"❌ Errores: {error}";
            CorrectoTextBlock.Text = $"✓ Correctos: {correcto}";

            // Habilitar/deshabilitar botón de corrección
            CorregirButton.IsEnabled = corregir > 0;
        }

        private void FilterButton_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;

            string filtro = btn.Tag as string;
            if (string.IsNullOrEmpty(filtro)) return;

            _filtroActual = filtro;
            AplicarFiltro();
        }

        private void AplicarFiltro()
        {
            List<DiagnosticRow> filasFiltradas;

            switch (_filtroActual)
            {
                case "Corregir":
                    filasFiltradas = _todasLasFilas.Where(r => r.Estado == EstadoParametro.Corregir).ToList();
                    break;
                case "Advertencia":
                    filasFiltradas = _todasLasFilas.Where(r => r.Estado == EstadoParametro.Advertencia).ToList();
                    break;
                case "Error":
                    filasFiltradas = _todasLasFilas.Where(r => r.Estado == EstadoParametro.Error).ToList();
                    break;
                case "Correcto":
                    filasFiltradas = _todasLasFilas.Where(r => r.Estado == EstadoParametro.Correcto).ToList();
                    break;
                case "Total":
                default:
                    filasFiltradas = _todasLasFilas;
                    break;
            }

            DiagnosticDataGrid.ItemsSource = filasFiltradas;
        }

        private void MostrarButton_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;

            DiagnosticRow row = btn.Tag as DiagnosticRow;
            if (row == null || row.ElementId == null || row.ElementId == ElementId.InvalidElementId)
                return;

            try
            {
                // Seleccionar elemento en Revit
                _uidoc.Selection.SetElementIds(new List<ElementId> { row.ElementId });
                _uidoc.ShowElements(row.ElementId);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"No se pudo mostrar el elemento:\n\n{ex.Message}");
            }
        }

        private void HeaderCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var row in _todasLasFilas)
            {
                row.IsChecked = true;
            }
        }

        private void HeaderCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (var row in _todasLasFilas)
            {
                row.IsChecked = false;
            }
        }

        private async void CorregirButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Contar elementos marcados para corregir
                int parametrosMarcados = _todasLasFilas.Count(r => r.IsChecked && r.Estado == EstadoParametro.Corregir);

                if (parametrosMarcados == 0)
                    return;

                // Confirmación
                var result = MessageBox.Show(
                    $"Se corregirán {parametrosMarcados} parámetros Uniclass marcados.\n\n¿Desea continuar?",
                    "Confirmar Corrección",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes)
                    return;

                // Filtrar solo los elementos marcados para corregir
                var filasSeleccionadas = _todasLasFilas.Where(r => r.IsChecked && r.Estado == EstadoParametro.Corregir).ToList();
                var elementosACorregir = _elementosData.Where(ed =>
                    filasSeleccionadas.Any(r => r.ElementId == ed.ElementId)).ToList();

                // Deshabilitar botón durante la corrección
                CorregirButton.IsEnabled = false;

                // Ejecutar corrección en thread secundario (evita congelar UI)
                var processingResult = await Task.Run(() =>
                    _parameterWriterAsync.UpdateParametersAsync(_doc, elementosACorregir));

                // Solo mostrar errores
                if (!processingResult.Exitoso)
                {
                    TaskDialog errorDialog = new TaskDialog("Error en Corrección")
                    {
                        MainInstruction = "No se pudieron corregir todos los parámetros",
                        MainContent = processingResult.Mensaje ?? "Error desconocido.",
                        ExpandedContent = processingResult.Errores.Any() ?
                            string.Join("\n", processingResult.Errores) : string.Empty
                    };
                    errorDialog.Show();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Error al corregir parámetros:\n\n{ex.Message}");
            }
            finally
            {
                // Rehabilitar botón
                CorregirButton.IsEnabled = true;
            }
        }

        private void ExportarButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                FileName = $"Diagnostico_Uniclass_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var exportService = new ExcelExportService();
                    var filasVisibles = DiagnosticDataGrid.ItemsSource as List<DiagnosticRow> ?? _todasLasFilas;
                    byte[] fileBytes = exportService.ExportToExcel(filasVisibles.ToList());
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
