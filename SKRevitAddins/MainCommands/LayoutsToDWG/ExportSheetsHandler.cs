using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.LayoutsToDWG
{
    public class ExportSheetsHandler : IExternalEventHandler
    {
        public IList<ElementId> ViewIds { get; set; }
        public string TargetPath { get; set; }
        public DWGExportOptions Options { get; set; }
        public bool OpenFolder { get; set; }
        public Action<bool> BusySetter { get; set; }
        public Action<int, int> ProgressReporter { get; set; }

        /// <summary>
        /// Mapping sheetId → file name prefix (ví dụ "A101_Title").
        /// </summary>
        public Dictionary<ElementId, string> FileNames { get; set; }

        /// <summary>
        /// Mapping sheetId → tên sheet set (ví dụ "Level 1 Plans").
        /// </summary>
        public Dictionary<ElementId, string> SubFolders { get; set; }

        public bool IsCancelled { get; set; } = false;

        public void Execute(UIApplication app)
        {
            try
            {
                BusySetter?.Invoke(true);
                var doc = app.ActiveUIDocument.Document;
                int total = ViewIds.Count, current = 0;

                foreach (var vid in ViewIds)
                {
                    if (IsCancelled) break;
                    var vs = doc.GetElement(vid) as ViewSheet;
                    if (vs == null) continue;

                    // Lấy prefix và subfolder
                    string prefix = FileNames.TryGetValue(vid, out var f) ? f : vs.SheetNumber;
                    string setName = SubFolders.TryGetValue(vid, out var sf) ? sf : "";
                    string outFolder = string.IsNullOrEmpty(setName)
                        ? TargetPath
                        : Path.Combine(TargetPath, LayerExportHelper.Sanitize(setName));

                    Directory.CreateDirectory(outFolder);

                    // Xuất trực tiếp với tên prefix
                    doc.Export(outFolder, prefix, new List<ElementId> { vid }, Options);

                    current++;
                    ProgressReporter?.Invoke(current, total);
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
                ProgressReporter?.Invoke(0, 1);
            }
        }

        public string GetName() => "Export Sheets Handler";
    }
}
