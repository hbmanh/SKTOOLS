// LayerExportImport.cs
// AutoCAD Add-in: Export/Import/Purge Layers via Excel
// Final complete code with sorted export, mm-formatted Lineweight, condensed UI

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
                AcadApp.ShowModalDialog(frm);
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
            // Form properties
            Text = "Layer Manager - Shinken Group®";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(360, 200);

            // Layout: 3 columns, 3 rows
            layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 3
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));  // header
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 10));  // spacer
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));  // buttons

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
            btnExport = new Button { Text = "Export...", Dock = DockStyle.Fill, Margin = new Padding(5) };
            btnImport = new Button { Text = "Import...", Dock = DockStyle.Fill, Margin = new Padding(5) };
            btnPurge = new Button { Text = "Purge All", Dock = DockStyle.Fill, Margin = new Padding(5) };
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
                        // Sort layer names alphabetically
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
                    // SaveAs with full params
                    object missing = Missing.Value;
                    wb.SaveAs(file, Excel.XlFileFormat.xlOpenXMLWorkbook,
                        missing, missing, false, false,
                        Excel.XlSaveAsAccessMode.xlNoChange,
                        missing, missing, missing, missing, missing);
                    xlApp.Visible = true;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Export lỗi:\n" + ex.Message);
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
                        .FirstOrDefault(w => w.FullName == file)
                        ?? xlApp.Workbooks.Open(file, false, true,
                            Missing.Value, Missing.Value, Missing.Value,
                            false, Missing.Value, Missing.Value,
                            Missing.Value, Missing.Value,
                            Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    if (wb.Worksheets.Count < 2) throw new Exception("Thiếu 2 sheet.");
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
                            if (!lt.Has(oldName))
                            {
                                errors.Add($"Dòng {i + 2}: '{oldName}' không tồn tại.");
                                continue;
                            }
                            string newName = b["Name"];
                            var recId = lt[oldName];
                            var ltr = (LayerTableRecord)tr.GetObject(recId, OpenMode.ForWrite);
                            // Rename if not layer "0"
                            if (!string.Equals(oldName, "0") && newName != oldName)
                            {
                                if (newName.IndexOfAny(Path.GetInvalidFileNameChars()) < 0)
                                {
                                    try
                                    {
                                        ltr.Name = newName;
                                        recId = lt[newName];
                                        ltr = (LayerTableRecord)tr.GetObject(recId, OpenMode.ForWrite);
                                    }
                                    catch (System.Exception ex)
                                    {
                                        errors.Add($"Dòng {i + 2}: rename lỗi {ex.Message}");
                                        continue;
                                    }
                                }
                                else errors.Add($"Dòng {i + 2}: tên không hợp lệ");
                            }
                            // Color
                            short origC = short.Parse(a["ColorIndex"]);
                            short newC = short.Parse(b["ColorIndex"]);
                            if (newC != origC)
                            {
                                try { ltr.Color = AcadColor.FromColorIndex(ColorMethod.ByAci, newC); }
                                catch (System.Exception ex) { errors.Add($"Dòng {i + 2}: Color lỗi {ex.Message}"); }
                            }
                            // LineWeight
                            string lwText = b["Lineweight"];
                            if (!string.IsNullOrWhiteSpace(lwText))
                            {
                                var part = lwText.Split(' ')[0];
                                if (double.TryParse(part, out double mmVal))
                                {
                                    int lwVal = (int)Math.Round(mmVal * 100);
                                    if (Enum.IsDefined(typeof(LineWeight), lwVal))
                                    {
                                        try { ltr.LineWeight = (LineWeight)lwVal; }
                                        catch (System.Exception ex) { errors.Add($"Dòng {i + 2}: LW lỗi {ex.Message}"); }
                                    }
                                    else errors.Add($"Dòng {i + 2}: LW '{lwText}' không hợp lệ");
                                }
                                else errors.Add($"Dòng {i + 2}: parse LW thất bại '{lwText}'");
                            }
                            // Linetype
                            string origLT = a["Linetype"], newLT = b["Linetype"];
                            if (!string.Equals(origLT, newLT, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    var ltt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);
                                    if (!ltt.Has(newLT)) db.LoadLineTypeFile(newLT, "acad.lin");
                                    if (ltt.Has(newLT)) ltr.LinetypeObjectId = ltt[newLT];
                                    else errors.Add($"Dòng {i + 2}: không tìm LT '{newLT}'");
                                }
                                catch (System.Exception ex) { errors.Add($"Dòng {i + 2}: LT lỗi {ex.Message}"); }
                            }
                            // PlotStyleName (optional)
                            string newPs = b["PlotStyleName"] ?? string.Empty;
                            if (!string.IsNullOrEmpty(newPs) && !string.Equals(newPs, ltr.PlotStyleName ?? string.Empty, StringComparison.Ordinal))
                            {
                                try
                                {
                                    ltr.PlotStyleName = newPs;
                                }
                                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                {
                                    // Nếu đang ở ColorDependent mode, bỏ qua
                                    if (ex.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.PlotStyleInColorDependentMode)
                                    {
                                        // CTB mode - ignore
                                    }
                                    else
                                    {
                                        errors.Add($"Dòng {i + 2}: PlotStyle lỗi: {ex.Message}");
                                    }
                                }
                            }
                        }
                        tr.Commit();
                    }
                    MessageBox.Show(errors.Count > 0 ? string.Join("\n", errors) : "Import OK", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Import lỗi:\n" + ex.Message);
                }
                finally { if (wb != null) Marshal.ReleaseComObject(wb); }
            }
        }

        private void OnPurge(object sender, EventArgs e)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            doc?.SendStringToExecute("_-PURGE _ALL\nY\n", true, false, false);
            MessageBox.Show("Purge all completed.");
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
