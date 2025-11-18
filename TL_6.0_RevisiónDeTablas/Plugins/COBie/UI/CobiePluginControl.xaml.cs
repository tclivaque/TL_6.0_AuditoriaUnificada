using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TL60_RevisionDeTablas.Models;
using TL60_RevisionDeTablas.Plugins.COBie.Services;

namespace TL60_RevisionDeTablas.Plugins.COBie.UI
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
            int vacios = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Vacio);
            int aCorregir = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Corregir);
            int errores = _originalDiagnosticRows.Count(r => r.Estado == EstadoParametro.Error);

            TotalTextBlock.Text = $"Total: {total}";
            CorregirTextBlock.Text = $"🔧 A Corregir: {aCorregir}";
            VacioTextBlock.Text = $"⚠ Vacío: {vacios}";
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

        private async void CorregirButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CorregirButton.IsEnabled = false;
                CorregirButton.Content = "Corrigiendo...";
                var writerAsync = _writerAsync;
                var doc = _doc;
                var dataToUpdate = _elementosData.Where(ed => ed.ParametrosActualizar.Any()).ToList();

                if (!dataToUpdate.Any())
                {
                    TaskDialog.Show("Información", "No hay parámetros marcados para corregir.", TaskDialogCommonButtons.Close);
                    CorregirButton.IsEnabled = true;
                    CorregirButton.Content = "Corregir";
                    return;
                }

                ProcessingResult writeResult = await Task.Run(() => writerAsync.WriteParametersAsync(doc, dataToUpdate));

                if (!writeResult.Exitoso)
                {
                    TaskDialog errorDialog = new TaskDialog("Error al Corregir")
                    {
                        MainInstruction = "No se pudieron corregir todos los parámetros",
                        MainContent = writeResult.Mensaje ?? "Error desconocido.",
                        ExpandedContent = writeResult.Errores.Any() ? string.Join("\n", writeResult.Errores) : string.Empty
                    };
                    errorDialog.Show();
                    CorregirButton.IsEnabled = true;
                    CorregirButton.Content = "Corregir";
                    return;
                }

                TaskDialog.Show("Éxito", "Los parámetros se han corregido correctamente.");
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
                case "Vacio": _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Vacio).ToList(); break;
                case "Corregir": _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Corregir).ToList(); break;
                case "Error": _diagnosticRows = _originalDiagnosticRows.Where(r => r.Estado == EstadoParametro.Error).ToList(); break;
                case "Total": default: _diagnosticRows = _originalDiagnosticRows; break;
            }
            DiagnosticDataGrid.ItemsSource = null;
            DiagnosticDataGrid.ItemsSource = _diagnosticRows;
        }
    }
}