using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    public class LayoutsToDWGRequestHandler : IExternalEventHandler
    {
        readonly LayoutsToDWGViewModel _vm;
        public LayoutsToDWGRequestHandler(LayoutsToDWGViewModel vm) => _vm = vm;
        public string GetName() => "LayoutsToDWGRequestHandler";

        public void Execute(UIApplication app)
        {
            var doc = app.ActiveUIDocument.Document;
            var folder = _vm.ExportFolder;
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            // DWG options
            var dwgOptions = DWGExportOptions
                .GetPredefinedOptions(doc, _vm.SelectedExportSetup)
                ?? new DWGExportOptions();

            // 1) Xuất tạm từng sheet
            _vm.PrepareSelectedSheets();
            var sheets = _vm.SelectedSheets;
            var tmpFolder = Path.Combine(folder, "tmp");
            Directory.CreateDirectory(tmpFolder);

            var tempFiles = new List<string>();
            foreach (var s in sheets)
            {
                string fname = $"{Path.GetFileNameWithoutExtension(_vm.MergedFilename)}-{s.SheetNumber}.dwg";
                doc.Export(tmpFolder, fname, new List<ElementId> { s.Id }, dwgOptions);
                tempFiles.Add(Path.Combine(tmpFolder, fname));
            }

            if (_vm.IsMergeMode)
            {
                // 2) Ghi JSON
                var cfg = new
                {
                    Files = tempFiles,
                    Output = Path.Combine(folder, _vm.MergedFilename),
                    MergeLayers = _vm.MergeLayers,
                    RasterToOle = _vm.RasterToOle
                };
                string jsonPath = Path.Combine(tmpFolder, "AcadCmd.json");
                File.WriteAllText(jsonPath, JsonConvert.SerializeObject(cfg, Formatting.Indented));

                // 3) Ghi script .scr
                string acadCmd = _vm.IsPremium ? "MSMERGEMLA" : "MSMERGE";
                File.WriteAllText(
                    AcadAppFolder.AcadCmdFile,
                    $"-SCRIPT\n\"{jsonPath}\"\n{acadCmd}\n-QSAVE\n"
                );

                // 4) Chạy AutoCAD merge
                string script = AcadAppFolder.AcadCmdFile;
                string output = Path.Combine(folder, _vm.MergedFilename);
                new RunAutoCadBg(
                    _vm.OpenAfterExport,  // openFile
                    script,               // scriptFile (.scr bạn đã tạo)
                    output,               // outputFile
                    silent: true          // chạy ẩn
                );
            }
            else
            {
                // Per-sheet: chỉ mở thư mục
                if (_vm.OpenAfterExport)
                    Process.Start("explorer.exe", folder);
            }
        }
    }
}
