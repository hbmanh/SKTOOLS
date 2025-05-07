using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;

namespace SKRevitAddins.LayoutsToDWG
{
    public static class LayerExportHelper
    {
        public static void WriteLayerMapping(Document doc, string setupName, string filePath)
        {
            var opt = DWGExportOptions.GetPredefinedOptions(doc, setupName)
                      ?? throw new InvalidOperationException($"Không tìm thấy Export Setup “{setupName}”.");

            var tbl = opt.GetExportLayerTable();
            if (tbl == null || !tbl.GetKeys().Any())
                throw new InvalidOperationException("Setup này không có layer‑mapping.");

            string[] skip = { ".dwg", ".dxf", ".nwc", ".nwd", ".ifc" };

            var sb = new StringBuilder(8192)
               .AppendLine("# shinken - Revit Export Layers")
               .AppendLine("# Category <> Subcategory <> Layer <> Color <> CutLayer <> CutColor")
               .AppendLine("# --------------------------------------------------------------");

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
            if (string.IsNullOrWhiteSpace(s)) return "Unnamed";
            char[] invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
                sb.Append(invalid.Contains(c) ? '_' : c);
            return sb.ToString().Trim().TrimEnd('.');
        }

        // Gộp SettingsHelper tại đây
        private const string FILE = "settings.json";
        private static string PathFile => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Shinken", "SheetsToDWG", FILE);

        public class ExportSettings
        {
            public string ExportPath { get; set; }
            public List<string> Params { get; set; }
            public List<string> Seps { get; set; }
        }

        public static ExportSettings LoadSettings()
        {
            try
            {
                if (File.Exists(PathFile))
                    return JsonSerializer.Deserialize<ExportSettings>(File.ReadAllText(PathFile));
            }
            catch { }
            return null;
        }

        public static void SaveSettings(ExportSettings s)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(PathFile)!);
            File.WriteAllText(PathFile,
                JsonSerializer.Serialize(s, new JsonSerializerOptions { WriteIndented = true }));
        }
    }
}
