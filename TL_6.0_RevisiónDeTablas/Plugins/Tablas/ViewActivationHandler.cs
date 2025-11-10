using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Threading;

namespace TL60_RevisionDeTablas.Services
{
    /// <summary>
    /// Maneja la activación de una vista (abrir tabla) usando ExternalEvent
    /// </summary>
    public class ViewActivationHandler : IExternalEventHandler
    {
        private ElementId _viewId = null;
        private ManualResetEvent _resetEvent;

        /// <summary>
        /// Configura los datos antes de ejecutar
        /// </summary>
        public void SetData(ElementId viewId, ManualResetEvent resetEvent)
        {
            _viewId = viewId;
            _resetEvent = resetEvent;
        }

        /// <summary>
        /// Se ejecuta en el contexto de Revit
        /// </summary>
        public void Execute(UIApplication app)
        {
            try
            {
                if (_viewId != null && _viewId != ElementId.InvalidElementId)
                {
                    UIDocument uidoc = app.ActiveUIDocument;
                    View viewToActivate = uidoc.Document.GetElement(_viewId) as View;

                    if (viewToActivate != null)
                    {
                        // ¡Acción principal! Cambia la vista activa.
                        uidoc.ActiveView = viewToActivate;
                    }
                }
            }
            catch (Exception)
            {
                // Ignorar errores, solo no abrirá la vista
            }
            finally
            {
                _resetEvent?.Set();
            }
        }

        public string GetName()
        {
            return "ViewActivationHandler";
        }
    }

    /// <summary>
    /// Clase auxiliar para manejar la activación de vistas asíncrona
    /// </summary>
    public class ViewActivatorAsync
    {
        private ExternalEvent _externalEvent;
        private ViewActivationHandler _handler;

        public ViewActivatorAsync()
        {
            _handler = new ViewActivationHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        /// <summary>
        /// Activa una vista de forma asíncrona
        /// </summary>
        public void ActivateViewAsync(ElementId viewId)
        {
            using (var resetEvent = new ManualResetEvent(false))
            {
                _handler.SetData(viewId, resetEvent);
                _externalEvent.Raise();
                resetEvent.WaitOne(TimeSpan.FromSeconds(10)); // Timeout 10s
            }
        }
    }
}