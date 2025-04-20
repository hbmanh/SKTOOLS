// LayerExportImport.cs
// AutoCAD Add-in: Async Export/Import/Purge Layers via Excel
// Updated: background processing with Task.Run to keep UI responsive

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Colors;

// Resolve ambiguous types
using SysImage = System.Drawing.Image;
using DrawColor = System.Drawing.Color;
using AcadColor = Autodesk.AutoCAD.Colors.Color;
using DrawFont = System.Drawing.Font;
using Exception = System.Exception;

[assembly: CommandClass(typeof(MyAutoCADAddin.LayerExportImport))]

namespace MyAutoCADAddin
{
    public class LayerExportImport : IExtensionApplication
    {
        public void Initialize() { }
        public void Terminate() { }

        [CommandMethod("lx")]
        public void Execute()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            using (var frm = new LayerExportImportForm())
                AcadApp.ShowModalDialog(frm);
        }
    }

    internal class LayerExportImportForm : Form
    {
        private Button btnExport, btnImport, btnPurge;
        private PictureBox logoBox;
        private Label lblTitle;
        private TableLayoutPanel layout;
        private ProgressBar progressBar;
        private const string LogoPath = @"C:\ProgramData\Autodesk\Revit\Addins\2023\SKTools.bundle\Contents\Resources\Images\shinken.png";

        public LayerExportImportForm()
        {
            // Form settings
            Text = "Layer Manager - Shinken Group®";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(360, 230);

            // Layout: columns and rows
            layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 4 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));  // header
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));  // progress
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 10));  // spacer
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));  // buttons

            // Logo
            logoBox = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom };
            if (File.Exists(LogoPath)) logoBox.Image = SysImage.FromFile(LogoPath);
            layout.Controls.Add(logoBox, 0, 0);

            // Title label
            lblTitle = new Label
            {
                Text = "Shinken Group®",
                Font = new DrawFont("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = DrawColor.FromArgb(45, 45, 48),
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };
            layout.Controls.Add(lblTitle, 1, 0);
            layout.SetColumnSpan(lblTitle, 2);

            // ProgressBar
            progressBar = new ProgressBar { Dock = DockStyle.Fill, Visible = false, Style = ProgressBarStyle.Marquee };
            layout.Controls.Add(progressBar, 0, 1);
            layout.SetColumnSpan(progressBar, 3);

            // Spacer (empty)
            layout.Controls.Add(new Label(), 0, 2);
            layout.SetColumnSpan(layout.GetControlFromPosition(0, 2) ?? new Control(), 3);

            // Buttons
            btnExport = new Button { Text = "Export...", Dock = DockStyle.Fill, Margin = new Padding(5) };
            btnImport = new Button { Text = "Import...", Dock = DockStyle.Fill, Margin = new Padding(5) };
            btnPurge = new Button { Text = "Purge All", Dock = DockStyle.Fill, Margin = new Padding(5) };
            btnExport.Click += OnExport;
            btnImport.Click += OnImport;
            btnPurge.Click += OnPurge;
            layout.Controls.Add(btnExport, 0, 3);
            layout.Controls.Add(btnImport, 1, 3);
            layout.Controls.Add(btnPurge, 2, 3);

            Controls.Add(layout);
        }
        private void OnPurge(object sender, EventArgs e)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.SendStringToExecute("_-PURGE _ALL\nY\n", true, false, false);
                MessageBox.Show("Purge all completed.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private async void OnExport(object sender, EventArgs e)
        {
            ToggleUI(false, "Exporting...");
            var file = AskSaveFile();
            if (file != null)
            {
                try
                {
                    await Task.Run(() => ExportLayersToExcel(file));
                    MessageBox.Show("Export completed!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Export error:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            ToggleUI(true);
        }

        private async void OnImport(object sender, EventArgs e)
        {
            ToggleUI(false, "Importing...");
            var file = AskOpenFile();
            if (file != null)
            {
                try
                {
                    var errors = await Task.Run(() => ImportLayersFromExcel(file));
                    if (errors.Any())
                        MessageBox.Show(string.Join("\n", errors), "Warnings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    else
                        MessageBox.Show("Import completed!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Import error:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            ToggleUI(true);
        }

        private void ToggleUI(bool enabled, string status = null)
        {
            btnExport.Enabled = btnImport.Enabled = btnPurge.Enabled = enabled;
            progressBar.Visible = !enabled;
            lblTitle.Text = enabled ? "Shinken Group®" : status;
        }

        private string AskSaveFile()
        {
            using (var sfd = new SaveFileDialog { Filter = "Excel Workbook|*.xlsx" })
                return sfd.ShowDialog() == DialogResult.OK ? sfd.FileName : null;
        }

        private string AskOpenFile()
        {
            using (var ofd = new OpenFileDialog { Filter = "Excel Workbook|*.xlsx;*.xls" })
                return ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : null;
        }

        // Core export logic
        private void ExportLayersToExcel(string excelFile)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws1 = null, ws2 = null;
            try
            {
                xlApp = new Excel.Application();
                wb = xlApp.Workbooks.Add();
                ws1 = (Excel.Worksheet)wb.Worksheets[1]; ws1.Name = "Layers";
                ws2 = (Excel.Worksheet)wb.Worksheets.Add(After: ws1); ws2.Name = "LayerRename";

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    int row = 1;
                    string[] headers = { "Name", "ColorIndex", "Linetype", "Lineweight", "PlotStyleName" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        ws1.Cells[row, i + 1] = headers[i];
                        ws2.Cells[row, i + 1] = headers[i];
                    }
                    row++;
                    var sorted = lt.Cast<ObjectId>()
                        .OrderBy(id => ((LayerTableRecord)tr.GetObject(id, OpenMode.ForRead)).Name);
                    foreach (var id in sorted)
                    {
                        var rec = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                        ws1.Cells[row, 1] = rec.Name;
                        ws1.Cells[row, 2] = rec.Color.ColorIndex;
                        string ltName = rec.LinetypeObjectId.IsNull
                            ? "ByLayer"
                            : ((LinetypeTableRecord)tr.GetObject(rec.LinetypeObjectId, OpenMode.ForRead)).Name;
                        ws1.Cells[row, 3] = ltName;
                        double lwMm = (int)rec.LineWeight / 100.0;
                        ws1.Cells[row, 4] = $"{lwMm:F2} mm";
                        ws1.Cells[row, 5] = rec.PlotStyleName ?? string.Empty;
                        for (int c = 1; c <= 5; c++)
                            ws2.Cells[row, c] = ws1.Cells[row, c].Value;
                        row++;
                    }
                    tr.Commit();
                }
                object missing = Missing.Value;
                wb.SaveAs(excelFile, Excel.XlFileFormat.xlOpenXMLWorkbook,
                    missing, missing, false, false,
                    Excel.XlSaveAsAccessMode.xlNoChange,
                    missing, missing, missing, missing, missing);
                xlApp.Visible = true;
            }
            finally
            {
                if (ws2 != null) Marshal.ReleaseComObject(ws2);
                if (ws1 != null) Marshal.ReleaseComObject(ws1);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        }

        // Core import logic; returns list of error messages
        private List<string> ImportLayersFromExcel(string file)
        {
            var errors = new List<string>();
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            Excel.Application xlApp = GetExcelFromROT();
            Excel.Workbook wb = xlApp.Workbooks.Cast<Excel.Workbook>()
                .FirstOrDefault(w => w.FullName == file)
                ?? xlApp.Workbooks.Open(file, false, true, Missing.Value, Missing.Value, Missing.Value,
                    false, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            var ws1 = (Excel.Worksheet)wb.Worksheets["Layers"];
            var ws2 = (Excel.Worksheet)wb.Worksheets["LayerRename"];
            var list1 = ReadSheet(ws1);
            var list2 = ReadSheet(ws2);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                for (int i = 0; i < list1.Count; i++)
                {
                    var a = list1[i];
                    var b = list2[i];
                    string oldName = a["Name"];
                    if (!lt.Has(oldName)) { errors.Add($"Dòng {i + 2}: '{oldName}' không tồn tại."); continue; }
                    string newName = b["Name"];
                    var recId = lt[oldName];
                    var ltr = (LayerTableRecord)tr.GetObject(recId, OpenMode.ForWrite);
                    // rename...
                    if (!string.Equals(oldName, "0") && newName != oldName)
                    {
                        if (newName.IndexOfAny(Path.GetInvalidFileNameChars()) < 0)
                        {
                            try { ltr.Name = newName; recId = lt[newName]; ltr = (LayerTableRecord)tr.GetObject(recId, OpenMode.ForWrite); }
                            catch (System.Exception ex) { errors.Add($"Dòng {i + 2}: rename lỗi {ex.Message}"); continue; }
                        }
                        else errors.Add($"Dòng {i + 2}: tên không hợp lệ");
                    }
                    // color
                    short origC = short.Parse(a["ColorIndex"]);
                    short newC = short.Parse(b["ColorIndex"]);
                    if (newC != origC)
                        TrySet(() => ltr.Color = AcadColor.FromColorIndex(ColorMethod.ByAci, newC), errors, i + 2, "Color");
                    // lineweight
                    var lwText = b["Lineweight"];
                    if (!string.IsNullOrWhiteSpace(lwText))
                    {
                        var part = lwText.Split(' ')[0];
                        if (double.TryParse(part, out double mmVal))
                        {
                            int lw = (int)Math.Round(mmVal * 100);
                            if (Enum.IsDefined(typeof(LineWeight), lw))
                                TrySet(() => ltr.LineWeight = (LineWeight)lw, errors, i + 2, "LineWeight");
                            else errors.Add($"Dòng {i + 2}: LW '{lwText}' không hợp lệ");
                        }
                        else errors.Add($"Dòng {i + 2}: parse LW thất bại '{lwText}'");
                    }
                    // linetype
                    var origLT = a["Linetype"];
                    var newLT = b["Linetype"];
                    if (!string.Equals(origLT, newLT, StringComparison.OrdinalIgnoreCase))
                    {
                        TrySet(() =>
                        {
                            var ltt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);
                            if (!ltt.Has(newLT)) db.LoadLineTypeFile(newLT, "acad.lin");
                            if (ltt.Has(newLT)) ltr.LinetypeObjectId = ltt[newLT];
                            else throw new Exception($"LT '{newLT}' not found");
                        }, errors, i + 2, "Linetype");
                    }
                    // plot style
                    var ps = b["PlotStyleName"];
                    if (!string.IsNullOrEmpty(ps) && !string.Equals(ps, ltr.PlotStyleName ?? string.Empty, StringComparison.Ordinal))
                    {
                        try { ltr.PlotStyleName = ps; }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            if (ex.ErrorStatus != Autodesk.AutoCAD.Runtime.ErrorStatus.PlotStyleInColorDependentMode)
                                errors.Add($"Dòng {i + 2}: PlotStyle lỗi {ex.Message}");
                        }
                    }
                }
                tr.Commit();
            }
            return errors;
        }

        // Helper to catch exceptions for property sets
        private void TrySet(Action action, List<string> errors, int row, string prop)
        {
            try { action(); }
            catch (Exception ex) { errors.Add($"Dòng {row}: {prop} lỗi {ex.Message}"); }
        }

        private List<Dictionary<string, string>> ReadSheet(Excel.Worksheet ws)
        {
            var data = new List<Dictionary<string, string>>();
            var used = ws.UsedRange;
            int rows = used.Rows.Count;
            for (int r = 2; r <= rows; r++)
            {
                var name = (ws.Cells[r, 1].Value ?? string.Empty).ToString();
                if (string.IsNullOrWhiteSpace(name)) break;
                data.Add(new Dictionary<string, string>
                {
                    ["Name"] = name,
                    ["ColorIndex"] = (ws.Cells[r, 2].Value ?? "0").ToString(),
                    ["Linetype"] = (ws.Cells[r, 3].Value ?? string.Empty).ToString(),
                    ["Lineweight"] = (ws.Cells[r, 4].Value ?? string.Empty).ToString(),
                    ["PlotStyleName"] = (ws.Cells[r, 5].Value ?? string.Empty).ToString()
                });
            }
            return data;
        }

        private Excel.Application GetExcelFromROT()
        {
            GetRunningObjectTable(0, out IRunningObjectTable rot);
            rot.EnumRunning(out IEnumMoniker enumMoniker);
            IMoniker[] mon = new IMoniker[1];
            while (enumMoniker.Next(1, mon, IntPtr.Zero) == 0)
            {
                CreateBindCtx(0, out IBindCtx ctx);
                mon[0].GetDisplayName(ctx, null, out string name);
                if (name.Contains("Excel"))
                {
                    rot.GetObject(mon[0], out object o);
                    return o as Excel.Application;
                }
            }
            return new Excel.Application();
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);
        [DllImport("ole32.dll")] private static extern int CreateBindCtx(int reserved, out IBindCtx ctx);
    }
}