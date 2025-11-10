// Plugins/Tablas/ScheduleWriterHandler.cs
using System;
using System.Collections.Generic;
using System.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TL60_RevisionDeTablas.Models; // <--- (¡CORRECCIÓN!) Añadido

namespace TL60_RevisionDeTablas.Plugins.Tablas // <--- (¡CORRECCIÓN!) Namespace actualizado
{
    /// <summary>
    /// Maneja la escritura de correcciones usando ExternalEvent
    /// </summary>
    public class ScheduleUpdateHandler : IExternalEventHandler
    {
        private Document _doc;
        private List<ElementData> _elementosData;
        private ProcessingResult _result;
        private ManualResetEvent _resetEvent;

        public ProcessingResult Result => _result;

        public void SetData(
            Document doc,
            List<ElementData> elementosData,
            ManualResetEvent resetEvent)
        {
            _doc = doc;
            _elementosData = elementosData;
            _resetEvent = resetEvent;
            _result = null;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                // (Ahora 'ScheduleUpdateWriter' está en el mismo namespace)
                var writer = new ScheduleUpdateWriter(_doc); // <--- (Línea 39)
                _result = writer.UpdateSchedules(_elementosData);
            }
            catch (Exception ex)
            {
                _result = new ProcessingResult
                {
                    Exitoso = false,
                    Mensaje = $"Error al escribir correcciones: {ex.Message}"
                };
                _result.Errores.Add(ex.Message);
            }
            finally
            {
                _resetEvent?.Set();
            }
        }

        public string GetName()
        {
            return "ScheduleUpdateHandler";
        }
    }

    /// <summary>
    /// Clase auxiliar para manejar la escritura asíncrona
    /// </summary>
    public class ScheduleUpdateAsync
    {
        private ExternalEvent _externalEvent;
        private ScheduleUpdateHandler _handler;

        public ScheduleUpdateAsync()
        {
            _handler = new ScheduleUpdateHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        /// <summary>
        /// Escribe correcciones de forma asíncrona
        /// </summary>
        public ProcessingResult UpdateSchedulesAsync(
            Document doc,
            List<ElementData> elementosData)
        {
            using (var resetEvent = new ManualResetEvent(false))
            {
                _handler.SetData(doc, elementosData, resetEvent);
                _externalEvent.Raise();

                // Timeout extendido a 600 segundos
                bool completed = resetEvent.WaitOne(TimeSpan.FromSeconds(600));

                if (!completed)
                {
                    return new ProcessingResult { Exitoso = false, Mensaje = "Timeout: La escritura tardó más de 10 minutos." };
                }

                return _handler.Result ?? new ProcessingResult { Exitoso = false, Mensaje = "Error: No se pudo obtener el resultado." };
            }
        }
    }
}