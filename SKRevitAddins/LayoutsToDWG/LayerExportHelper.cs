using Autodesk.Revit.DB;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace SKRevitAddins.LayoutsToDWG
{
    public static class LayerExportHelper
    {
        public static void WriteLayerMapping(Document doc, string setupName, string filePath)
        {
            var opt = DWGExportOptions.GetPredefinedOptions(doc, setupName)
                      ?? throw new InvalidOperationException(
                             $"Không tìm thấy Export Setup “{setupName}”.");

            var tbl = opt.GetExportLayerTable();
            if (tbl == null || !tbl.GetKeys().Any())
                throw new InvalidOperationException("Setup này không có layer‑mapping.");

            string[] skip = { ".dwg", ".dxf", ".nwc", ".nwd", ".ifc" };

            var sb = new StringBuilder(8192)
               .AppendLine("# shinken - Revit Export Layers")
               .AppendLine("# Category <> Subcategory <> Layer <> Color <> CutLayer <> CutColor")
               .AppendLine("# -----------------------------------------------------");

            foreach (var k in tbl.GetKeys())
            {
                if (skip.Any(e => k.CategoryName.IndexOf(e, StringComparison.OrdinalIgnoreCase) >= 0))
                    continue;

                var info = tbl[k];
                if (info.ColorNumber < 0) continue;

                sb.Append(k.CategoryName).Append('\t')
                  .Append(k.SubCategoryName ?? "").Append('\t')
                  .Append(info.LayerName).Append('\t')
                  .Append(info.ColorNumber);

                if (!string.IsNullOrWhiteSpace(info.CutLayerName))
                    sb.Append('\t').Append(info.CutLayerName)
                      .Append('\t').Append(info.CutColorNumber);
                sb.AppendLine();
            }

            Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        public static string Sanitize(string s)
        {
            foreach (char c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
            return s.Trim();
        }
    }
}
