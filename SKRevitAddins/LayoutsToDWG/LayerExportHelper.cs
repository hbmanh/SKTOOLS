using Autodesk.Revit.DB;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;

namespace SKRevitAddins.LayoutsToDWG
{
    public static class LayerExportHelper
    {
        //────────────────────────────────────────────────────────
        // 1. EXPORT LAYER‑MAPPING
        //────────────────────────────────────────────────────────
        public static void ExportLayerMappingToTxt(
            Document doc, string dwgSettingName, string exportFilePath)
        {
            DWGExportOptions opt =
                DWGExportOptions.GetPredefinedOptions(doc, dwgSettingName)
                ?? throw new InvalidOperationException(
                     $"Không tìm thấy DWGExportSettings “{dwgSettingName}”.");

            ExportLayerTable tbl = opt.GetExportLayerTable();
            if (tbl == null || !tbl.GetKeys().Any())
                throw new InvalidOperationException("Setup này không có layer‑mapping.");

            // 🔍 các phần mở rộng CAD/IFC cần bỏ qua
            string[] skipExts = { ".dwg", ".dxf", ".nwc", ".nwd", ".ifc" };

            var sb = new StringBuilder(8192);
            sb.AppendLine("# MLabs - Revit Export Layers");
            sb.AppendLine("# Maps Categories and Subcategories to layer names and color numbers");
            sb.AppendLine("# Category <> Subcategory <> Layer name <> Color number <> Cut layer name <> Cut color number");
            sb.AppendLine("# -----------------------------------------------------");

            foreach (ExportLayerKey k in tbl.GetKeys())
            {
                string cat = k.CategoryName;

                // ❶ BỎ layer link/import
                if (skipExts.Any(ext =>
                        cat.IndexOf(ext, StringComparison.OrdinalIgnoreCase) >= 0))
                    continue;

                ExportLayerInfo info = tbl[k];

                // ❷ Bỏ dòng chưa định màu
                if (info.ColorNumber < 0) continue;

                // ❸ Ghi 4 cột bắt buộc
                sb.Append(cat).Append('\t')
                  .Append(k.SubCategoryName ?? string.Empty).Append('\t')
                  .Append(info.LayerName).Append('\t')
                  .Append(info.ColorNumber);

                // ❹ Ghi thêm 2 cột Cut nếu có
                if (!string.IsNullOrWhiteSpace(info.CutLayerName))
                {
                    sb.Append('\t').Append(info.CutLayerName)
                      .Append('\t').Append(info.CutColorNumber); // có thể =‑1
                }
                sb.AppendLine();
            }

            Directory.CreateDirectory(Path.GetDirectoryName(exportFilePath)!);
            File.WriteAllText(exportFilePath, sb.ToString(), Encoding.UTF8);
        }

        //────────────────────────────────────────────────────────
        // 2. VALIDATE FILE
        //────────────────────────────────────────────────────────
        public static bool IsValidLayerMappingFile(string path)
        {
            if (!File.Exists(path)) return false;

            try
            {
                foreach (string raw in File.ReadLines(path))
                {
                    if (string.IsNullOrWhiteSpace(raw) || raw.StartsWith("#"))
                        continue;

                    string[] cols = raw.Split('\t');
                    if (!(cols.Length == 4 || cols.Length == 6)) return false;

                    // Cột 4: màu chính 0‑255
                    if (!byte.TryParse(cols[3], out _)) return false;

                    // Nếu có Cut: cột 6 phải 0‑255 hoặc ‑1
                    if (cols.Length == 6 &&
                        !(int.TryParse(cols[5], out int cut) &&
                          (cut == -1 || (cut >= 0 && cut <= 255))))
                        return false;
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
