using Autodesk.Revit.DB;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace SKRevitAddins.LayoutsToDWG
{
    public static class LayerExportHelper
    {
        public static void ExportLayerMappingToTxt(Document doc, string dwgSettingName, string exportFilePath)
        {
            DWGExportOptions options = DWGExportOptions.GetPredefinedOptions(doc, dwgSettingName);

            if (options == null)
                throw new InvalidOperationException($"DWGExportSettings '{dwgSettingName}' not found.");

            ExportLayerTable layerTable = options.GetExportLayerTable();
            if (layerTable == null || !layerTable.GetKeys().Any())
                throw new InvalidOperationException("No layer mapping found.");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("# Exported DWG Layer Mapping");
            sb.AppendLine("# Format: Category\tSubCategory\tLayer\tColor\tCutLayer\tCutColor");

            foreach (ExportLayerKey key in layerTable.GetKeys())
            {
                ExportLayerInfo info = layerTable[key];

                string category = key.CategoryName;
                string subCategory = string.IsNullOrWhiteSpace(key.SubCategoryName) ? "-" : key.SubCategoryName;
                string layer = info.LayerName ?? "-";
                string color = info.ColorNumber >= 0 ? info.ColorNumber.ToString() : "-";
                string cutLayer = string.IsNullOrWhiteSpace(info.CutLayerName) ? "-" : info.CutLayerName;
                string cutColor = info.CutColorNumber >= 0 ? info.CutColorNumber.ToString() : "-";

                sb.AppendLine($"{category}\t{subCategory}\t{layer}\t{color}\t{cutLayer}\t{cutColor}");
            }

            File.WriteAllText(exportFilePath, sb.ToString());
        }
        public static bool IsValidLayerMappingFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    return false;

                var lines = File.ReadAllLines(filePath)
                    .Where(line => !string.IsNullOrWhiteSpace(line) && !line.StartsWith("#"))
                    .ToList();

                foreach (var line in lines)
                {
                    var parts = line.Split('\t');
                    if (parts.Length < 6)
                        return false;

                    // Optional: Kiểm tra dữ liệu định dạng từng cột nếu cần
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

    }
}