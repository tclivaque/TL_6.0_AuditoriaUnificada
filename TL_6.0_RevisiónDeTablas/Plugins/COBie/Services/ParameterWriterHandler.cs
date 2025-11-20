using System;
using System.Collections.Generic;
using System.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TL60_AuditoriaUnificada.Plugins.COBie.Models;
using TL60_AuditoriaUnificada.Models;

namespace TL60_AuditoriaUnificada.Plugins.COBie.Services
{
    public class ParameterWriterHandler : IExternalEventHandler
    {
        private Document _doc;
        private List<ElementData> _elementosData;
        private ProcessingResult _result;
        private ManualResetEvent _resetEvent;
        private bool _shouldRepaint;
        private List<string> _categorias;
        // Campos _config y _definitions eliminados

        public ProcessingResult Result => _result;

        // SetData vuelve a su firma original
        public void SetData(Document doc, List<ElementData> elementosData, ManualResetEvent resetEvent,
                            bool shouldRepaint = false, List<string> categorias = null)
        {
            _doc = doc;
            _elementosData = elementosData;
            _resetEvent = resetEvent;
            _shouldRepaint = shouldRepaint;
            _categorias = categorias;
            _result = null;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                // ===== CAMBIO AQUÍ: Usar constructor original de ParameterWriter =====
                var writer = new ParameterWriter(_doc);
                _result = writer.WriteParameters(_elementosData);

                // La lógica de repintado permanece igual
                if (_shouldRepaint)
                {
                    try
                    {
                        var viewManager = new ViewGraphicsManager(_doc, _doc.ActiveView, _categorias);
                        if (viewManager.IsView3D())
                        {
                            viewManager.ApplyGraphicsAndIsolate(_elementosData);
                        }
                    }
                    catch (Exception ex)
                    {
                        _result.Errores.Add($"Error al repintar modelo 3D: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                _result = new ProcessingResult { Exitoso = false, Mensaje = $"Error al escribir parámetros: {ex.Message}" };
                _result.Errores.Add(ex.Message);
            }
            finally
            {
                _resetEvent?.Set();
            }
        }

        // Métodos GetCobieConfigFromSomewhere y GetDefinitionsFromSomewhere ELIMINADOS

        public string GetName()
        {
            return "ParameterWriterHandler";
        }
    }

    public class ParameterWriterAsync
    {
        private ExternalEvent _externalEvent;
        private ParameterWriterHandler _handler;
        // Campos _config y _definitions eliminados
        // Método Initialize eliminado

        public ParameterWriterAsync()
        {
            _handler = new ParameterWriterHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        public ProcessingResult WriteParametersAsync(Document doc, List<ElementData> elementosData)
        {
            using (var resetEvent = new ManualResetEvent(false))
            {
                // ===== CAMBIO AQUÍ: Llamada a SetData original =====
                _handler.SetData(doc, elementosData, resetEvent,
                                 shouldRepaint: false, categorias: null);
                _externalEvent.Raise();

                // Usamos 900 segundos (15 min) que pidió el usuario
                bool completed = resetEvent.WaitOne(TimeSpan.FromSeconds(900));

                if (!completed)
                {
                    return new ProcessingResult { Exitoso = false, Mensaje = "Timeout: La escritura tardó más de 15 minutos." };
                }
                return _handler.Result ?? new ProcessingResult { Exitoso = false, Mensaje = "Error: No se pudo obtener resultado." };
            }
        }

        // WriteParametersAndRepaint también usa la firma original de SetData
        public ProcessingResult WriteParametersAndRepaint(Document doc, List<ElementData> elementosData, List<string> categorias)
        {
            using (var resetEvent = new ManualResetEvent(false))
            {
                // ===== CAMBIO AQUÍ: Llamada a SetData original =====
                _handler.SetData(doc, elementosData, resetEvent,
                                 shouldRepaint: true, categorias: categorias);
                _externalEvent.Raise();

                bool completed = resetEvent.WaitOne(TimeSpan.FromSeconds(900)); // 15 min timeout

                if (!completed)
                {
                    return new ProcessingResult { Exitoso = false, Mensaje = "Timeout: La operación tardó más de 15 minutos." };
                }
                return _handler.Result ?? new ProcessingResult { Exitoso = false, Mensaje = "Error: No se pudo obtener resultado." };
            }
        }
    }
}