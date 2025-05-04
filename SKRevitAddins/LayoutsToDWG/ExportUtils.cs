using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace SKRevitAddins.LayoutsToDWG
{
    /// <summary>Tiện ích lấy danh sách các cấu hình DWG Export trong file Revit.</summary>
    internal static class ExportUtils
    {
        public static IList<ExportDWGSettings> GetExportDWGSettings(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ExportDWGSettings))
                .Cast<ExportDWGSettings>()
                .ToList();
        }
    }
}