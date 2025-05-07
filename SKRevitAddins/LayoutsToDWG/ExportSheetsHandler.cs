// SKRevitAddins.LayoutsToDWG.ExportSheetsHandler.cs
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SKRevitAddins.LayoutsToDWG
{
    public class ExportSheetsHandler : IExternalEventHandler
    {
        public IList<ElementId> ViewIds { get; set; }
        public string TargetPath { get; set; }
        public DWGExportOptions Options { get; set; }
        public bool OpenFolder { get; set; }

        public Action<bool> BusySetter { get; set; }

        // NEW: Dictionary mapping ViewId → filename
        public Dictionary<ElementId, string> FileNames { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                BusySetter?.Invoke(true);
                var doc = app.ActiveUIDocument.Document;

                // Export từng sheet
                foreach (var vid in ViewIds)
                {
                    doc.Export(TargetPath, "", new List<ElementId> { vid }, Options);
                }

                // Đổi tên file
                foreach (var id in ViewIds)
                {
                    var vs = doc.GetElement(id) as ViewSheet;
                    if (vs == null) continue;

                    string dstCore = FileNames.TryGetValue(id, out var name)
                        ? name : vs.SheetNumber;

                    string srcDwg = Directory.GetFiles(TargetPath, $"*{vs.SheetNumber}*.dwg")
                        .FirstOrDefault(f => Path.GetFileName(f)
                            .IndexOf("View -", StringComparison.OrdinalIgnoreCase) < 0);

                    if (srcDwg != null)
                    {
                        string dstDwg = Path.Combine(TargetPath, dstCore + ".dwg");
                        if (File.Exists(dstDwg)) File.Delete(dstDwg);
                        File.Move(srcDwg, dstDwg);
                    }
                }

                if (OpenFolder)
                    Process.Start("explorer.exe", TargetPath);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("DWG Export", $"Error: {ex.Message}");
            }
            finally
            {
                BusySetter?.Invoke(false);
            }
        }

        public string GetName() => "Export Sheets Handler";
    }
}
