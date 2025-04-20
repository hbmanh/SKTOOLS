// LayerExportImport.cs
// AutoCAD Add-in: Export/Import/Purge Layers via Excel
// Final complete code

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
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
            {
                AcadApp.ShowModalDialog(frm);
            }
        }
    }

    internal class LayerExportImportForm : Form
    {
        private Button btnExport, btnImport, btnPurge;
        private PictureBox logoBox;
        private Label lblTitle;
        private TableLayoutPanel layout;
        private const string LogoPath = @"C:\ProgramData\Autodesk\Revit\Addins\2023\SKTools.bundle\Contents\Resources\Images\shinken.png";

        public LayerExportImportForm()
        {
            // Form settings
            Text = "Layer Manager - Shinken Group®";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(360, 180);

            // Layout: 3 rows, 3 columns
            layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 3
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));  // Header
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 10));  // Spacer
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));  // Buttons

            // Logo
            logoBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom
            };
            if (File.Exists(LogoPath))
                logoBox.Image = SysImage.FromFile(LogoPath);
            layout.Controls.Add(logoBox, 0, 0);

            // Title
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

            // Buttons
            btnExport = new Button { Text = "Export...", Size = new Size(80, 24), Margin = new Padding(3) };
            btnImport = new Button { Text = "Import...", Size = new Size(80, 24), Margin = new Padding(3) };
            btnPurge = new Button { Text = "Purge All", Size = new Size(80, 24), Margin = new Padding(3) };
            btnExport.Click += OnExport;
            btnImport.Click += OnImport;
            btnPurge.Click += OnPurge;
            layout.Controls.Add(btnExport, 0, 2);
            layout.Controls.Add(btnImport, 1, 2);
            layout.Controls.Add(btnPurge, 2, 2);

            Controls.Add(layout);
        }

        private void OnExport(object sender, EventArgs e)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            string defaultFile = Path.Combine(Path.GetDirectoryName(doc.Name), Path.GetFileNameWithoutExtension(doc.Name) + "_Layers.xlsx");
            using (var sfd = new SaveFileDialog { Filter = "Excel Workbook|*.xlsx", FileName = defaultFile, InitialDirectory = Path.GetDirectoryName(doc.Name) })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;
                string file = sfd.FileName;

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
                        // Sort by Name
                        var sorted = lt.Cast<ObjectId>()
                            .OrderBy(o => ((LayerTableRecord)tr.GetObject(o, OpenMode.ForRead)).Name)
                            .ToList();
                        foreach (ObjectId layerId in sorted)
                        {
                            var rec = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                            ws1.Cells[row, 1] = rec.Name;
                            ws1.Cells[row, 2] = rec.Color.ColorIndex;
                            ws1.Cells[row, 3] = rec.LinetypeObjectId.IsNull
                                ? "ByLayer"
                                : ((LinetypeTableRecord)tr.GetObject(rec.LinetypeObjectId, OpenMode.ForRead)).Name;
                            double lwMm = (int)rec.LineWeight / 100.0;
                            ws1.Cells[row, 4] = $"{lwMm:F2} mm";
                            ws1.Cells[row, 5] = rec.PlotStyleName ?? string.Empty;
                            for (int c = 1; c <= 5; c++)
                                ws2.Cells[row, c] = ws1.Cells[row, c].Value;
                            row++;
                        }
                        tr.Commit();
                    }
                    // SaveAs
                    object missing = Missing.Value;
                    wb.SaveAs(file,
                        Excel.XlFileFormat.xlOpenXMLWorkbook,
                        missing, missing, false, false,
                        Excel.XlSaveAsAccessMode.xlNoChange,
                        missing, missing, missing, missing, missing);
                    xlApp.Visible = true;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Export lỗi:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (ws2 != null) Marshal.ReleaseComObject(ws2);
                    if (ws1 != null) Marshal.ReleaseComObject(ws1);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            }
        }

        private void OnImport(object sender, EventArgs e)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            using (var ofd = new OpenFileDialog { Filter = "Excel Workbook|*.xlsx;*.xls", InitialDirectory = Path.GetDirectoryName(doc.Name) })
            {
                if (ofd.ShowDialog() != DialogResult.OK) return;
                string file = ofd.FileName;
                Excel.Application xlApp = null;
                Excel.Workbook wb = null;
                var errors = new List<string>();
                try
                {
                    xlApp = GetExcelFromROT();
                    wb = xlApp.Workbooks.Cast<Excel.Workbook>()
                        .FirstOrDefault(wb2 => string.Equals(wb2.FullName, file, StringComparison.OrdinalIgnoreCase))
                        ?? xlApp.Workbooks.Open(file, false, true,
                            Missing.Value, Missing.Value, Missing.Value,
                            false, Missing.Value, Missing.Value,
                            Missing.Value, Missing.Value,
                            Missing.Value, Missing.Value,
                            Missing.Value, Missing.Value);
                    if (wb.Worksheets.Count < 2) throw new Exception("File phải có 2 Sheet.");
                    var ws1 = (Excel.Worksheet)wb.Worksheets["Layers"];
                    var ws2 = (Excel.Worksheet)wb.Worksheets["LayerRename"];
                    if (ws1.UsedRange.Columns.Count != 5) throw new Exception("Sai số cột.");
                    var list1 = ReadSheet(ws1);
                    var list2 = ReadSheet(ws2);
                    if (list1.Count != list2.Count) throw new Exception("Sai số dòng.");

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                        for (int i = 0; i < list1.Count; i++)
                        {
                            var a = list1[i];
                            var b = list2[i];
                            string oldName = a["Name"];
                            if (!lt.Has(oldName)) { errors.Add($"Dòng {i + 2}: Layer '{oldName}' không tồn tại."); continue; }
                            var recId = lt[oldName];
                            var ltr = (LayerTableRecord)tr.GetObject(recId, OpenMode.ForWrite);

                            // Rename except '0'
                            string newName = b["Name"];
                            if (!string.Equals(oldName, "0", StringComparison.OrdinalIgnoreCase)
                                && !string.Equals(newName, oldName, StringComparison.Ordinal))
                            {
                                if (newName.IndexOfAny(Path.GetInvalidFileNameChars()) < 0)
                                {
                                    try { ltr.Name = newName; recId = lt[newName]; ltr = (LayerTableRecord)tr.GetObject(recId, OpenMode.ForWrite); }
                                    catch (System.Exception ex) { errors.Add($"Dòng {i + 2}: Rename lỗi '{ex.Message}'"); continue; }
                                }
                                else errors.Add($"Dòng {i + 2}: Tên '{newName}' không hợp lệ.");
                            }

                            // Color
                            short origCol = short.Parse(a["ColorIndex"]);
                            short newCol = short.Parse(b["ColorIndex"]);
                            if (newCol != origCol)
                            {
                                try { ltr.Color = AcadColor.FromColorIndex(ColorMethod.ByAci, newCol); }
                                catch (System.Exception ex) { errors.Add($"Dòng {i + 2}: Color lỗi '{ex.Message}'"); }
                            }

                            // LineWeight
                            string lwText = b["Lineweight"];
                            if (!string.IsNullOrWhiteSpace(lwText))
                            {
                                var parts = lwText.Split(' ');
                                if (double.TryParse(parts[0], out double mmVal))
                                {
                                    int lwVal = (int)Math.Round(mmVal * 100);
                                    if (Enum.IsDefined(typeof(LineWeight), lwVal)
                                        && (int)ltr.LineWeight != lwVal)
                                    {
                                        try { ltr.LineWeight = (LineWeight)lwVal; }
                                        catch (System.Exception ex) { errors.Add($"Dòng {i + 2}: LineWeight lỗi '{ex.Message}'"); }
                                    }
                                }
                                else errors.Add($"Dòng {i + 2}: Không parse Lineweight '{lwText}'");
                            }

                            // Linetype
                            string origLt = a["Linetype"], newLt = b["Linetype"];
                            if (!string.Equals(origLt, newLt, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    var ltt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);
                                    if (!ltt.Has(newLt)) db.LoadLineTypeFile(newLt, "acad.lin");
                                    if (ltt.Has(newLt)) ltr.LinetypeObjectId = ltt[newLt];
                                    else errors.Add($"Dòng {i + 2}: Không tìm thấy Linetype '{newLt}'");
                                }
                                catch (System.Exception ex) { errors.Add($"Dòng {i + 2}: Linetype lỗi '{ex.Message}'"); }
                            }

                            // PlotStyleName (optional)
                            string origPS = a["PlotStyleName"];
                            string newPS = b["PlotStyleName"];
                            if (newPS != origPS) ltr.PlotStyleName = newPS;
                        }
                        tr.Commit();
                    }

                    MessageBox.Show(errors.Count > 0
                        ? string.Join("\n", errors)
                        : "Import & cập nhật thành công!",
                        "Kết quả",
                        MessageBoxButtons.OK,
                        errors.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Import lỗi:\n" + ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            }
        }

        private void OnPurge(object sender, EventArgs e)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            doc?.SendStringToExecute("_-PURGE _ALL\nY\n", true, false, false);
            MessageBox.Show("Purge all completed.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private List<Dictionary<string, string>> ReadSheet(Excel.Worksheet ws)
        {
            var data = new List<Dictionary<string, string>>();
            var used = ws.UsedRange; int rows = used.Rows.Count;
            for (int r = 2; r <= rows; r++)
            {
                var name = (ws.Cells[r, 1].Value ?? "").ToString();
                if (string.IsNullOrWhiteSpace(name)) break;
                data.Add(new Dictionary<string, string>
                {
                    ["Name"] = name,
                    ["ColorIndex"] = (ws.Cells[r, 2].Value ?? "0").ToString(),
                    ["Linetype"] = (ws.Cells[r, 3].Value ?? "").ToString(),
                    ["Lineweight"] = (ws.Cells[r, 4].Value ?? "").ToString(),
                    ["PlotStyleName"] = (ws.Cells[r, 5].Value ?? "").ToString()
                });
            }
            return data;
        }

        private Excel.Application GetExcelFromROT()
        {
            GetRunningObjectTable(0, out IRunningObjectTable rot);
            rot.EnumRunning(out IEnumMoniker enumMoniker);
            IMoniker[] monikers = new IMoniker[1];
            while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
            {
                CreateBindCtx(0, out IBindCtx bindCtx);
                monikers[0].GetDisplayName(bindCtx, null, out string name);
                if (name.Contains("Excel"))
                {
                    rot.GetObject(monikers[0], out object comObj);
                    return comObj as Excel.Application;
                }
            }
            return new Excel.Application();
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);
        [DllImport("ole32.dll")] private static extern int CreateBindCtx(int reserved, out IBindCtx ctx);
    }
}
