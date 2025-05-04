using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace SKRevitAddins.LayoutsToDWG
{
    /// <summary>ExternalEvent Handler: xuất DWG hoặc Layer Mapping.</summary>
    public sealed class LayoutsToDWGRequestHandler : IExternalEventHandler
    {
        internal LayoutsToDWGViewModel? ViewModel { get; set; }

        public void Execute(UIApplication uiApp)
        {
            if (ViewModel == null) return;

            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            var req = ViewModel.CreateRequest();

            // --- 1. xuất Layer Mapping nếu người dùng chọn ---
            if (!string.IsNullOrWhiteSpace(req.LayerTxtPath))
            {
                ExportLayerTxt(doc, req);
                ViewModel.ResetLayerTxt();      // tránh lặp
                return;                         // kết thúc
            }

            // --- 2. xuất DWG (per‑sheet) ---------------------
            var opts = ExportUtils.GetExportDWGSettings(doc)
                       .FirstOrDefault(s => s.Name == req.ExportSetupName)?
                       .GetDWGExportOptions()
                       ?? new DWGExportOptions();

            foreach (var id in req.SheetIds)
            {
                var sheet = (ViewSheet)doc.GetElement(id);
                string num = GetSheetParam(sheet, req.SheetNumberParam) ?? sheet.SheetNumber;
                string nam = GetSheetParam(sheet, req.SheetNameParam) ?? sheet.Name;
                string fn = BuildFileName(req.Prefix, num, nam);

                doc.Export(req.ExportFolder, fn, new[] { id }, opts);
            }

            if (req.OpenAfterExport)
                Process.Start("explorer.exe", req.ExportFolder);
        }

        public string GetName() => "LayoutsToDWG – ExternalEvent Handler";

        // ---------- helpers ----------
        private static string? GetSheetParam(ViewSheet s, string? paramName) =>
            string.IsNullOrWhiteSpace(paramName) ? null :
            s.LookupParameter(paramName)?.AsString();

        private static string BuildFileName(string prefix, string number, string name)
        {
            string safe(string s) =>
                string.Concat(s.Split(Path.GetInvalidFileNameChars()));

            number = safe(number);
            name = safe(name);

            return string.IsNullOrWhiteSpace(prefix)
                   ? $"{number}_{name}"
                   : $"{prefix}-{number}_{name}";
        }

        private static void ExportLayerTxt(Document doc, LayoutsToDWGRequest req)
        {
            var settings = ExportUtils.GetExportDWGSettings(doc)
                           .FirstOrDefault(s => s.Name == req.ExportSetupName);
            if (settings == null) return;

            var table = settings.GetDWGExportOptions().GetExportLayerTable();
            var keys = table.GetKeys();

            using var sw = new StreamWriter(req.LayerTxtPath!, false, Encoding.UTF8);

            sw.WriteLine("# Revit Export Layers");
            sw.WriteLine("# Maps Categories and Subcategories to layer names and color numbers");
            sw.WriteLine("# Category <tab> Subcategory <tab> Layer name <tab> Color number <tab>");
            sw.WriteLine("# Cut layer name <tab> Cut color number");
            sw.WriteLine("# -----------------------------------------------------");

            foreach (var k in keys)
            {
                ExportLayerInfo info = table[k];

                string cat = k.CategoryName ?? "";
                string sub = k.SubCategoryName ?? "";
                string lay = info.LayerName ?? "";
                int col = info.ColorNumber;
                string layCut = info.CutLayerName ?? lay;
                int cutCol = info.CutColorNumber;

                sw.WriteLine($"{cat}\t{sub}\t{lay}\t{col}\t{layCut}\t{cutCol}");
            }
        }
    }
    internal sealed class LayoutsToDWGRequest
    {
        public IList<ElementId> SheetIds { get; set; } = new List<ElementId>();   // null khi chỉ export layer map

        public string? ExportSetupName { get; set; }
        public string ExportFolder { get; set; } = "";

        // ---- cấu trúc tên file DWG ----
        public string Prefix { get; set; } = "";
        public string? SheetNumberParam { get; set; }
        public string? SheetNameParam { get; set; }

        // ---- xuất layer mapping (.txt) ----
        public string? LayerTxtPath { get; set; }      // null  ⇒ không export

        public bool OpenAfterExport { get; set; }
    }
}
