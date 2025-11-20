// Plugins/Uniclass/UI/UniclassPluginControl.xaml.cs
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TL60_AuditoriaUnificada.Models;
using TL60_AuditoriaUnificada.Plugins.Uniclass.Services;

namespace TL60_AuditoriaUnificada.Plugins.Uniclass.UI
{
    public partial class UniclassPluginControl : UserControl
    {
        private readonly UIDocument _uidoc;
        private readonly Document _doc;
        private List<DiagnosticRow> _todasLasFilas;
        private List<ElementData> _elementosData;
        private string _filtroActual = "Total";

        public UniclassPluginControl(UIDocument uidoc, List<ElementData> elementosData)
        {
            InitializeComponent();
            _uidoc = uidoc ?? throw new ArgumentNullException(nameof(uidoc));
            _doc = uidoc.Document;
            _elementosData = elementosData ?? new List<ElementData>();
            _todasLasFilas = new List<DiagnosticRow>();

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
                MessageBox.Show($"Error al cargar datos de Uniclass:\n\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ActualizarContadores()
        {
            int total = _todasLasFilas.Count;
            int corregir = _todasLasFilas.Count(r => r.Estado == EstadoParametro.Corregir);
            int vacio = _todasLasFilas.Count(r => r.Estado == EstadoParametro.Vacio);
            int error = _todasLasFilas.Count(r => r.Estado == EstadoParametro.Error);
            int correcto = _todasLasFilas.Count(r => r.Estado == EstadoParametro.Correcto);

            TotalTextBlock.Text = $"Total: {total}";
            CorregirTextBlock.Text = $"🔧 A Corregir: {corregir}";
            VacioTextBlock.Text = $"⚠ Advertencias: {vacio}";
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
                case "Vacio":
                    filasFiltradas = _todasLasFilas.Where(r => r.Estado == EstadoParametro.Vacio).ToList();
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
                MessageBox.Show($"No se pudo mostrar el elemento:\n\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void CorregirButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Contar elementos a corregir
                int parametrosACorregir = _todasLasFilas.Count(r => r.Estado == EstadoParametro.Corregir);

                if (parametrosACorregir == 0)
                {
                    MessageBox.Show("No hay parámetros para corregir.",
                        "Información", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Confirmación
                var result = MessageBox.Show(
                    $"Se corregirán {parametrosACorregir} parámetros Uniclass.\n\n¿Desea continuar?",
                    "Confirmar Corrección",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes)
                    return;

                // Ejecutar corrección
                var writer = new UniclassParameterWriter(_doc);
                var processingResult = writer.UpdateParameters(_elementosData);

                // Mostrar resultado
                if (processingResult.Exitoso)
                {
                    MessageBox.Show(processingResult.Mensaje,
                        "Corrección Exitosa",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    // Recargar datos
                    RecargarDatos();
                }
                else
                {
                    MessageBox.Show($"Se encontraron errores durante la corrección:\n\n{processingResult.Mensaje}",
                        "Errores en Corrección",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al corregir parámetros:\n\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RecargarDatos()
        {
            try
            {
                // Nota: Después de corregir, necesitamos volver a procesar los elementos
                // porque los valores actuales han cambiado
                MessageBox.Show(
                    "Los parámetros han sido corregidos exitosamente.\n\n" +
                    "Por favor, cierre y vuelva a abrir la ventana de auditoría para ver los cambios.",
                    "Recarga Requerida",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al recargar datos:\n\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
