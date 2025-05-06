using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SKRevitAddins.LayoutsToDWG
{
    public class LayerMappingEntry
    {
        public string Category { get; set; }
        public string SubCategory { get; set; }
        public string Layer { get; set; }
        public int Color { get; set; } = -1;
        public string CutLayer { get; set; }
        public int CutColor { get; set; } = -1;
    }

    public static class LayerMappingParser
    {
        public static bool IsValidLayerMappingFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return false;

                var lines = File.ReadAllLines(filePath)
                    .Where(line => !string.IsNullOrWhiteSpace(line) && !line.StartsWith("#"));

                return lines.All(line => line.Split('\t').Length >= 6);
            }
            catch
            {
                return false;
            }
        }

        public static List<LayerMappingEntry> Parse(string filePath)
        {
            var entries = new List<LayerMappingEntry>();

            var lines = File.ReadAllLines(filePath)
                .Where(line => !string.IsNullOrWhiteSpace(line) && !line.StartsWith("#"));

            foreach (var line in lines)
            {
                var parts = line.Split('\t');
                if (parts.Length < 6) continue;

                entries.Add(new LayerMappingEntry
                {
                    Category = parts[0],
                    SubCategory = parts[1] == "-" ? "" : parts[1],
                    Layer = parts[2],
                    Color = int.TryParse(parts[3], out var c) ? c : -1,
                    CutLayer = parts[4] == "-" ? "" : parts[4],
                    CutColor = int.TryParse(parts[5], out var cc) ? cc : -1
                });
            }

            return entries;
        }

        public static void WriteToFile(string filePath, IEnumerable<LayerMappingEntry> entries)
        {
            using var writer = new StreamWriter(filePath);
            writer.WriteLine("# Exported DWG Layer Mapping");
            writer.WriteLine("# Format: Category\tSubCategory\tLayer\tColor\tCutLayer\tCutColor");

            foreach (var e in entries)
            {
                writer.WriteLine($"{e.Category}\t{(string.IsNullOrWhiteSpace(e.SubCategory) ? "-" : e.SubCategory)}\t{e.Layer}\t{(e.Color >= 0 ? e.Color.ToString() : "-")}\t{(string.IsNullOrWhiteSpace(e.CutLayer) ? "-" : e.CutLayer)}\t{(e.CutColor >= 0 ? e.CutColor.ToString() : "-")}");
            }
        }
    }
}
