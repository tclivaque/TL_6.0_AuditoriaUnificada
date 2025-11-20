using Autodesk.Revit.DB;

namespace TL60_AuditoriaUnificada.Models
{
    /// <summary>
    /// Almacena la definición de un filtro correcto
    /// </summary>
    public class ScheduleFilterInfo
    {
        public string FieldName { get; set; }
        public ScheduleFilterType FilterType { get; set; }
        public object Value { get; set; }

        public string AsString()
        {
            return $"{FieldName} {FilterType} {Value}";
        }
    }
}