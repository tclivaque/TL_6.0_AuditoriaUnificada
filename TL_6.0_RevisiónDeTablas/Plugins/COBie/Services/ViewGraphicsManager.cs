using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using TL60_AuditoriaUnificada.Plugins.COBie.Models;
using TL60_AuditoriaUnificada.Models;

namespace TL60_AuditoriaUnificada.Plugins.COBie.Services
{
    /// <summary>
    /// Administrador de gráficos de vista (pintar y aislar elementos)
    /// </summary>
    public class ViewGraphicsManager
    {
        private readonly Document _doc;
        private readonly View _activeView;
        private readonly List<string> _categorias;

        public ViewGraphicsManager(Document doc, View activeView, List<string> categorias = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _activeView = activeView;
            _categorias = categorias ?? new List<string>(); // Si es null, lista vacía
        }

        /// <summary>
        /// Verifica si la vista actual es 3D
        /// </summary>
        public bool IsView3D()
        {
            return _activeView is View3D;
        }

        /// <summary>
        /// Aplica colores a elementos según su estado y los aisla
        /// </summary>
        public void ApplyGraphicsAndIsolate(List<ElementData> elementosData)
        {
            if (!IsView3D())
            {
                // No hacer nada si no es vista 3D (continuar silenciosamente)
                return;
            }

            using (Transaction trans = new Transaction(_doc, "Aplicar Colores COBie"))
            {
                trans.Start();

                try
                {
                    // Clasificar elementos por estado
                    var elementosCorrectos = new List<ElementId>();
                    var elementosACorregir = new List<ElementId>();
                    var elementosVacios = new List<ElementId>();
                    var elementosError = new List<ElementId>();

                    foreach (var elementData in elementosData)
                    {
                        if (elementData.Element == null) continue;

                        ElementId id = elementData.Element.Id;

                        // Determinar estado del elemento
                        if (elementData.GrupoCOBie == "ERROR")
                        {
                            elementosError.Add(id);
                        }
                        else if (elementData.ParametrosVacios.Count > 0)
                        {
                            elementosVacios.Add(id);
                        }
                        else if (elementData.ParametrosActualizar.Count > 0)
                        {
                            elementosACorregir.Add(id);
                        }
                        else if (elementData.CobieCompleto)
                        {
                            elementosCorrectos.Add(id);
                        }
                    }

                    // Obtener TODOS los elementos visibles en la vista
                    var todosElementosVisibles = GetAllVisibleElements();

                    // Elementos procesados = todos los que tienen color específico
                    var elementosProcesados = new List<ElementId>();
                    elementosProcesados.AddRange(elementosCorrectos);
                    elementosProcesados.AddRange(elementosACorregir);
                    elementosProcesados.AddRange(elementosVacios);
                    elementosProcesados.AddRange(elementosError);

                    // Elementos sin procesar = resto
                    var elementosSinProcesar = todosElementosVisibles
                        .Where(id => !elementosProcesados.Contains(id))
                        .ToList();

                    // Crear overrides
                    var overrideVerde = CreateOverrideVerde();      // Weight 1, Halftone False
                    var overrideAzul = CreateOverrideAzul();        // Weight 12, Halftone False
                    var overrideAmarillo = CreateOverrideAmarillo(); // Weight 12, Halftone False
                    var overrideRojo = CreateOverrideRojo();        // Weight 12, Halftone False
                    var overrideGris = CreateOverrideGris();        // Weight 1, Halftone True

                    // Aplicar overrides
                    foreach (var id in elementosCorrectos)
                    {
                        try { _activeView.SetElementOverrides(id, overrideVerde); }
                        catch { continue; }
                    }

                    foreach (var id in elementosACorregir)
                    {
                        try { _activeView.SetElementOverrides(id, overrideAzul); }
                        catch { continue; }
                    }

                    foreach (var id in elementosVacios)
                    {
                        try { _activeView.SetElementOverrides(id, overrideAmarillo); }
                        catch { continue; }
                    }

                    foreach (var id in elementosError)
                    {
                        try { _activeView.SetElementOverrides(id, overrideRojo); }
                        catch { continue; }
                    }

                    foreach (var id in elementosSinProcesar)
                    {
                        try { _activeView.SetElementOverrides(id, overrideGris); }
                        catch { continue; }
                    }

                    // Aislar solo elementos procesados (con color)
                    if (elementosProcesados.Count > 0)
                    {
                        _activeView.IsolateElementsTemporary(elementosProcesados);
                    }

                    trans.Commit();
                }
                catch (Exception)
                {
                    trans.RollBack();
                    throw new Exception("Error al aplicar gráficos en vista 3D");
                }
            }
        }

        /// <summary>
        /// Crea override para elementos CORRECTOS (Verde - Weight 1 - Halftone False)
        /// </summary>
        private OverrideGraphicSettings CreateOverrideVerde()
        {
            var settings = new OverrideGraphicSettings();

            Color colorLineas = new Color(0, 200, 0);
            Color colorRelleno = new Color(128, 255, 128);

            ElementId patronSolidoFill = GetSolidFillPatternId();
            ElementId patronSolidoLine = LinePatternElement.GetSolidPatternId();

            settings.SetProjectionLineColor(colorLineas);
            settings.SetProjectionLineWeight(1);
            settings.SetProjectionLinePatternId(patronSolidoLine);

            settings.SetSurfaceForegroundPatternColor(colorRelleno);
            if (patronSolidoFill != null && patronSolidoFill != ElementId.InvalidElementId)
            {
                settings.SetSurfaceForegroundPatternId(patronSolidoFill);
            }

            settings.SetHalftone(false);

            return settings;
        }

        /// <summary>
        /// Crea override para elementos A CORREGIR (Azul - Weight 12 - Halftone False)
        /// </summary>
        private OverrideGraphicSettings CreateOverrideAzul()
        {
            var settings = new OverrideGraphicSettings();

            Color colorLineas = new Color(66, 139, 202);    // Azul medio
            Color colorRelleno = new Color(209, 236, 241);  // Azul claro

            ElementId patronSolidoFill = GetSolidFillPatternId();
            ElementId patronSolidoLine = LinePatternElement.GetSolidPatternId();

            settings.SetProjectionLineColor(colorLineas);
            settings.SetProjectionLineWeight(12);
            settings.SetProjectionLinePatternId(patronSolidoLine);

            settings.SetSurfaceForegroundPatternColor(colorRelleno);
            if (patronSolidoFill != null && patronSolidoFill != ElementId.InvalidElementId)
            {
                settings.SetSurfaceForegroundPatternId(patronSolidoFill);
            }

            settings.SetHalftone(false);

            return settings;
        }

        /// <summary>
        /// Crea override para elementos VACÍOS (Amarillo - Weight 12 - Halftone False)
        /// </summary>
        private OverrideGraphicSettings CreateOverrideAmarillo()
        {
            var settings = new OverrideGraphicSettings();

            Color colorLineas = new Color(255, 193, 7);     // Amarillo oscuro
            Color colorRelleno = new Color(255, 243, 205);  // Amarillo claro

            ElementId patronSolidoFill = GetSolidFillPatternId();
            ElementId patronSolidoLine = LinePatternElement.GetSolidPatternId();

            settings.SetProjectionLineColor(colorLineas);
            settings.SetProjectionLineWeight(12);
            settings.SetProjectionLinePatternId(patronSolidoLine);

            settings.SetSurfaceForegroundPatternColor(colorRelleno);
            if (patronSolidoFill != null && patronSolidoFill != ElementId.InvalidElementId)
            {
                settings.SetSurfaceForegroundPatternId(patronSolidoFill);
            }

            settings.SetHalftone(false);

            return settings;
        }

        /// <summary>
        /// Crea override para elementos ERROR (Rojo - Weight 12 - Halftone False)
        /// </summary>
        private OverrideGraphicSettings CreateOverrideRojo()
        {
            var settings = new OverrideGraphicSettings();

            Color colorLineas = new Color(210, 0, 0);
            Color colorRelleno = new Color(248, 215, 218);

            ElementId patronSolidoFill = GetSolidFillPatternId();
            ElementId patronSolidoLine = LinePatternElement.GetSolidPatternId();

            settings.SetProjectionLineColor(colorLineas);
            settings.SetProjectionLineWeight(12);
            settings.SetProjectionLinePatternId(patronSolidoLine);

            settings.SetSurfaceForegroundPatternColor(colorRelleno);
            if (patronSolidoFill != null && patronSolidoFill != ElementId.InvalidElementId)
            {
                settings.SetSurfaceForegroundPatternId(patronSolidoFill);
            }

            settings.SetHalftone(false);

            return settings;
        }

        /// <summary>
        /// Crea override para elementos SIN PROCESAR (Gris - Weight 1 - Halftone True)
        /// </summary>
        private OverrideGraphicSettings CreateOverrideGris()
        {
            var settings = new OverrideGraphicSettings();

            Color colorLineas = new Color(128, 128, 128);
            Color colorRelleno = new Color(255, 255, 255);

            ElementId patronSolidoFill = GetSolidFillPatternId();
            ElementId patronSolidoLine = LinePatternElement.GetSolidPatternId();

            settings.SetProjectionLineColor(colorLineas);
            settings.SetProjectionLineWeight(1);
            settings.SetProjectionLinePatternId(patronSolidoLine);

            settings.SetSurfaceForegroundPatternColor(colorRelleno);
            if (patronSolidoFill != null && patronSolidoFill != ElementId.InvalidElementId)
            {
                settings.SetSurfaceForegroundPatternId(patronSolidoFill);
            }

            settings.SetHalftone(true); // Solo grises tienen Halftone

            return settings;
        }

        /// <summary>
        /// Obtiene el patrón sólido para Fill
        /// </summary>
        private ElementId GetSolidFillPatternId()
        {
            try
            {
                var allPatterns = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FillPatternElement))
                    .Cast<FillPatternElement>()
                    .ToList();

                foreach (var pattern in allPatterns)
                {
                    if (pattern.GetFillPattern().IsSolidFill)
                    {
                        return pattern.Id;
                    }
                }

                return allPatterns.FirstOrDefault()?.Id ?? ElementId.InvalidElementId;
            }
            catch
            {
                return ElementId.InvalidElementId;
            }
        }

        /// <summary>
        /// Obtiene todos los elementos visibles en la vista actual
        /// AHORA USA LAS CATEGORÍAS DESDE _categorias (leídas de Google Sheets)
        /// </summary>
        private List<ElementId> GetAllVisibleElements()
        {
            var elementos = new List<ElementId>();

            // Si no hay categorías configuradas, usar lista por defecto (fallback)
            var categoriasAProcesar = _categorias != null && _categorias.Count > 0
                ? _categorias
                : GetDefaultCategories();

            foreach (var categoriaString in categoriasAProcesar)
            {
                try
                {
                    // Parsear string a BuiltInCategory
                    if (Enum.TryParse<BuiltInCategory>(categoriaString, out BuiltInCategory categoria))
                    {
                        var collector = new FilteredElementCollector(_doc, _activeView.Id)
                            .OfCategory(categoria)
                            .WhereElementIsNotElementType();

                        elementos.AddRange(collector.ToElementIds());
                    }
                }
                catch
                {
                    // Si falla al procesar una categoría, continuar con la siguiente
                    continue;
                }
            }

            return elementos;
        }

        /// <summary>
        /// Lista por defecto de categorías (fallback si no se leen desde Sheets)
        /// </summary>
        private List<string> GetDefaultCategories()
        {
            return new List<string>
    {
        // MEP
        "OST_PlumbingFixtures",
        "OST_PipeFitting",
        "OST_PipeCurves",
        "OST_PipeAccessory",
        "OST_FlexPipeCurves",
        "OST_ConduitFitting",
        "OST_Conduit",
        "OST_ElectricalEquipment",
        "OST_ElectricalFixtures",
        "OST_DataDevices",
        "OST_SecurityDevices",
        "OST_FireAlarmDevices",
        "OST_CommunicationDevices",
        "OST_NurseCallDevices",
        "OST_LightingFixtures",
        
        // Arquitectura
        "OST_StructuralColumns",
        "OST_StructuralFraming",
        "OST_Floors",
        "OST_Roofs",
        "OST_Walls",
        "OST_Doors",
        "OST_Windows",
        "OST_GenericModel"
    };
        }

        /// <summary>
        /// Limpia los overrides y deshace el aislamiento
        /// </summary>
        public void ClearGraphicsAndUnisolate()
        {
            if (!IsView3D())
            {
                return;
            }

            using (Transaction trans = new Transaction(_doc, "Limpiar Colores COBie"))
            {
                trans.Start();

                try
                {
                    // Deshabilitar aislamiento temporal
                    _activeView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);

                    // Limpiar todos los overrides
                    var todosElementos = GetAllVisibleElements();
                    var overrideDefault = new OverrideGraphicSettings();

                    foreach (var elementId in todosElementos)
                    {
                        try
                        {
                            _activeView.SetElementOverrides(elementId, overrideDefault);
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    throw new Exception($"Error al limpiar gráficos: {ex.Message}", ex);
                }
            }
        }
    }
}