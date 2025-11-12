// Plugins/Fixer/FixLegacySchedulesCommand.cs
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using TL60_RevisionDeTablas.Models; // Reutilizaremos los modelos si es necesario

namespace TL60_RevisionDeTablas.Plugins.Fixer
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FixLegacySchedulesCommand : IExternalCommand
    {
        // Regex para encontrar el formato: [C.XX.XX] - [Descripción] - RNG
        // Grupo 1: (C\.\S+)       -> Captura el Assembly Code (ej. "C.20.07.05.09")
        // Grupo 2: (.*)           -> Captura la Descripción (ej. "CORTINA ACUSTICA...")
        // \s*-\s*RNG$             -> Busca el sufijo " - RNG" al final
        private static readonly Regex _nameRegex = new Regex(@"^(C\.\S+)\s*-\s*(.*)\s*-\s*RNG$", RegexOptions.Compiled);
        private const string EMPRESA_FILTER_NAME = "EMPRESA";

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            int namesFixed = 0;
            int filtersFixed = 0;

            // 1. Identificar solo las tablas de "Metrado"
            // (Nombre empieza con "C." Y GRUPO DE VISTA empieza con "C.")
            List<ViewSchedule> metradoSchedules;
            try
            {
                metradoSchedules = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Schedules)
                    .WhereElementIsNotElementType()
                    .Cast<ViewSchedule>()
                    .Where(v => v.Name.StartsWith("C.") &&
                                (v.LookupParameter("GRUPO DE VISTA")?.AsString() ?? "")
                                .StartsWith("C.", StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }
            catch (Exception ex)
            {
                message = $"Error al recolectar tablas: {ex.Message}";
                return Result.Failed;
            }

            if (metradoSchedules.Count == 0)
            {
                TaskDialog.Show("Información", "No se encontraron tablas de metrado (C./C.) para corregir.");
                return Result.Succeeded;
            }

            using (Transaction trans = new Transaction(doc, "Revertir Nombres y Filtros de Tablas"))
            {
                try
                {
                    trans.Start();

                    foreach (ViewSchedule view in metradoSchedules)
                    {
                        ScheduleDefinition definition = view.Definition;
                        bool nameChanged = false;
                        bool filterChanged = false;

                        // =============================================
                        // Tarea 1: Corregir Nombres
                        // =============================================
                        string oldName = view.Name;
                        Match match = _nameRegex.Match(oldName);

                        if (match.Success)
                        {
                            string ac = match.Groups[1].Value;
                            string desc = match.Groups[2].Value.Trim(); // Quitar espacios extra
                            string newName = $"{ac} {desc}"; // Formato [AC] [Desc]

                            if (view.Name != newName)
                            {
                                view.Name = newName;
                                namesFixed++;
                                nameChanged = true;
                            }
                        }

                        // =============================================
                        // Tarea 2: Eliminar Filtro "EMPRESA"
                        // =============================================
                        var currentFilters = definition.GetFilters();
                        var filtersToKeep = new List<ScheduleFilter>();
                        bool empresaFilterFound = false;

                        foreach (ScheduleFilter filter in currentFilters)
                        {
                            ScheduleField field = definition.GetField(filter.FieldId);
                            if (field == null) continue;

                            if (field.GetName().Equals(EMPRESA_FILTER_NAME, StringComparison.OrdinalIgnoreCase))
                            {
                                empresaFilterFound = true;
                            }
                            else
                            {
                                filtersToKeep.Add(filter);
                            }
                        }

                        // Si encontramos y quitamos el filtro, reconstruimos la lista
                        if (empresaFilterFound)
                        {
                            definition.ClearFilters();
                            foreach (ScheduleFilter filter in filtersToKeep)
                            {
                                definition.AddFilter(filter);
                            }
                            filtersFixed++;
                            filterChanged = true;
                        }
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    message = $"Error durante la transacción: {ex.Message}";
                    return Result.Failed;
                }
            }

            TaskDialog.Show("Corrección Completa",
                $"Proceso finalizado.\n\n" +
                $"- Nombres de tablas revertidos: {namesFixed}\n" +
                $"- Filtros 'EMPRESA' eliminados: {filtersFixed}");

            return Result.Succeeded;
        }
    }
}