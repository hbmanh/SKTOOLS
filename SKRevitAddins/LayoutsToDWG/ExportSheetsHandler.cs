using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Concurrent;
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
        public string FilePattern { get; set; }   // {num} {name}
        public bool OpenFolder { get; set; }
        public Action<bool> BusySetter { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                BusySetter?.Invoke(true);

                var doc = app.ActiveUIDocument.Document;
                var sw = Stopwatch.StartNew();

                // 1) Export một lần
                doc.Export(TargetPath, "", ViewIds, Options);

                // 2) Rename song song & log
                var log = new ConcurrentBag<string>();

                Parallel.ForEach(ViewIds, id =>
                {
                    var vs = (ViewSheet)doc.GetElement(id);
                    string num = vs.SheetNumber;
                    string name = vs.Name;

                    // tên gốc do Revit tạo
                    string src = Path.Combine(TargetPath, $"{num}.dwg");
                    if (!File.Exists(src))
                        src = Directory.GetFiles(TargetPath, $"{num}*.dwg").FirstOrDefault();

                    string dst = Path.Combine(TargetPath,
                                   LayerExportHelper.Sanitize(
                                        FilePattern.Replace("{num}", num)
                                                   .Replace("{name}", name)) + ".dwg");

                    if (src != null && File.Exists(src))
                    {
                        if (File.Exists(dst)) File.Delete(dst);   // *** .NET 4.x không hỗ trợ overwrite
                        File.Move(src, dst);
                    }
                    log.Add($"{num},{dst}");
                });

                File.WriteAllLines(Path.Combine(TargetPath, "_ExportLog.csv"),
                                   log.OrderBy(l => l));

                sw.Stop();
                TaskDialog.Show("DWG Export",
                    $"Xuất thành công {ViewIds.Count} sheet.\nThời gian: {sw.Elapsed:mm\\:ss}.");

                if (OpenFolder)
                    Process.Start("explorer.exe", TargetPath);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
            finally { BusySetter?.Invoke(false); }
        }

        public string GetName() => "Export Sheets Handler";
    }
}
