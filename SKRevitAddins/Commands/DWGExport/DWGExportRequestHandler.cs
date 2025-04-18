using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace SKRevitAddins.Commands.DWGExport
{
    public class DWGExportRequestHandler : IExternalEventHandler
    {
        private readonly DWGExportViewModel _vm;
        private readonly DWGExportRequest _req;

        public DWGExportRequestHandler(DWGExportViewModel vm, DWGExportRequest req)
        {
            _vm = vm;
            _req = req;
        }

        public string GetName() => "DWGExportRequestHandler";
        public DWGExportRequest Request => _req;

        public void Execute(UIApplication uiApp)
        {
            if (_req.Take() != LayerExportRequestId.Export) return;
            var doc = uiApp.ActiveUIDocument.Document;
            Export(doc);
        }

        private void Export(Document doc)
        {
            if (!_vm.SelectedSheets.Any())
            {
                _vm.ExportStatusMessage = "Please select sheet(s).";
                return;
            }

            using (var dlg = new FolderBrowserDialog { Description = "Folder for DWG" })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                int total = _vm.SelectedSheets.Count;
                _vm.ProgressMax = total;
                _vm.ProgressValue = 0;

                int i = 0;
                foreach (var si in _vm.SelectedSheets)
                {
                    i++;
                    var sheet = si.Sheet;
                    if (sheet == null) continue;

                    var options = BuildOptions(doc);
                    string fn = Sanitize($"{sheet.SheetNumber}-{sheet.Name}") + ".dwg";

                    try
                    {
                        doc.Export(dlg.SelectedPath, fn,
                                   new List<ElementId> { sheet.Id }, options);
                        _vm.ExportStatusMessage = $"Exported {fn}";
                    }
                    catch (System.Exception ex)
                    {
                        _vm.ExportStatusMessage = $"Error {fn}: {ex.Message}";
                    }

                    _vm.ProgressValue = i;
                }

                _vm.ExportStatusMessage = "Export complete!";
            }
        }

        private DWGExportOptions BuildOptions(Document doc)
        {
            var opt = new DWGExportOptions
            {
                MergedViews = false,
                TargetUnit = ExportUnit.Millimeter
            };
            DisableExternalRefs(opt);

            var asm = typeof(DWGExportOptions).Assembly;
            var tblType = asm.GetType("Autodesk.Revit.DB.ExportLayerTable");
            if (tblType != null)
            {
                object tbl = System.Activator.CreateInstance(tblType);
                var setName = tblType.GetMethod("SetLayerNameForCategory");
                var setColor = tblType.GetMethod("SetLayerColorForCategory");
                var propMapping = typeof(DWGExportOptions).GetProperty("LayerMapping");

                foreach (var row in _vm.CategoryItems)
                {
                    Category cat = doc.Settings.Categories.get_Item(row.CategoryName);
                    if (cat == null) continue;

                    setName.Invoke(tbl, new object[] { cat.Id, row.LayerName });
                    setColor.Invoke(tbl, new object[]
                    {
                        cat.Id,
                        ColorUtil.AciToRevit(row.ColorIndex)
                    });
                }

                if (propMapping != null)
                    propMapping.SetValue(opt, tbl, null);
            }

            return opt;
        }

        private static void DisableExternalRefs(DWGExportOptions o)
        {
            foreach (var p in typeof(DWGExportOptions).GetProperties())
            {
                if (p.CanWrite && p.PropertyType == typeof(bool) &&
                    p.Name.IndexOf("ExternalReference",
                                   System.StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    p.SetValue(o, false, null);
                }
            }
        }

        private static string Sanitize(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }
}
