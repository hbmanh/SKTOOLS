// SKRevitAddins.Commands.LayoutsToDWG.Helpers.cs
using Microsoft.Win32;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace SKRevitAddins.Commands.LayoutsToDWG.Helpers
{
    #region POCO cho JSON
    public class AcadMergeCommand
    {
        public List<string> SheetFiles { get; set; }
        public string FilePath { get; set; }
        public bool MergeLayers { get; set; } = true;
        public bool OpenFile2 { get; set; } = true;
    }
    #endregion

    #region Tìm AutoCAD
    public static class AutoCADLocator
    {
        public static string FindLowestAcadExe()
        {
            var list = new List<(int Year, string Path)>();
            for (int year = 2026; year >= 2013; year--)
            {
                string reg = $@"SOFTWARE\Autodesk\AutoCAD\R{year - 1998}";
                using var k = Registry.LocalMachine.OpenSubKey(reg) ?? Registry.CurrentUser.OpenSubKey(reg);
                if (k == null) continue;
                foreach (var sub in k.GetSubKeyNames())
                {
                    using var sk = k.OpenSubKey(sub);
                    string loc = sk?.GetValue("AcadLocation") as string;
                    if (string.IsNullOrEmpty(loc)) continue;
                    string exe = Path.Combine(loc, "acad.exe");
                    if (File.Exists(exe)) list.Add((year, exe));
                }
            }
            return list.OrderBy(p => p.Year).FirstOrDefault().Path
                ?? Enumerable.Range(2013, 14).Reverse()
                   .Select(y => $@"C:\Program Files\Autodesk\AutoCAD {y}\acad.exe")
                   .FirstOrDefault(File.Exists);
        }
    }
    #endregion

    #region Xuất JSON + SCR + chạy AutoCAD
    public static class DWGMergeHelper
    {
        public static void GenerateMergeCommandJson(string jsonPath, IEnumerable<string> dwgFiles,
                                                    string outputPath, bool mergeLayers, bool openAfter)
        {
            var cmd = new AcadMergeCommand
            {
                SheetFiles = dwgFiles.ToList(),
                FilePath = outputPath,
                MergeLayers = mergeLayers,
                OpenFile2 = openAfter
            };
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(cmd, Formatting.Indented));
        }

        public static void GenerateMergeCommandScript(string scrPath)
            => File.WriteAllText(scrPath, "MLabsMergeSheets\n");

        public static void RunAutoCADMerge(string acadExe, string scrPath)
        {
            if (!File.Exists(acadExe) || !File.Exists(scrPath)) return;
            var psi = new ProcessStartInfo
            {
                FileName = acadExe,
                Arguments = $"/nologo /b \"{scrPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Minimized
            };
            Process.Start(psi);
        }
    }
    #endregion
}
