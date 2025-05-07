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
        public Dictionary<ElementId, string> FileNames { get; set; }

        public bool IsCancelled { get; set; } = false;
        public Action<int, int> ProgressReporter { get; set; } // current, total

        public void Execute(UIApplication app)
        {
            try
            {
                BusySetter?.Invoke(true);
                var doc = app.ActiveUIDocument.Document;

                int total = ViewIds.Count;
                int current = 0;

                foreach (var vid in ViewIds)
                {
                    if (IsCancelled) break;

                    doc.Export(TargetPath, "", new List<ElementId> { vid }, Options);
                    current++;
                    ProgressReporter?.Invoke(current, total);
                }

                foreach (var id in ViewIds)
                {
                    if (IsCancelled) break;

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

                if (!IsCancelled && OpenFolder)
                    Process.Start("explorer.exe", TargetPath);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("DWG Export", $"Error: {ex.Message}");
            }
            finally
            {
                BusySetter?.Invoke(false);
                IsCancelled = false;
                ProgressReporter?.Invoke(0, 1); // reset progress
            }
        }

        public string GetName() => "Export Sheets Handler";
    }

}
