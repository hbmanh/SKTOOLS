using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Colors;
using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AutoColor = System.Drawing.Color;
using Color = Autodesk.AutoCAD.Colors.Color;
using Exception = Autodesk.AutoCAD.Runtime.Exception;
using FormsApp = System.Windows.Forms.Application;

[assembly: CommandClass(typeof(SKAcadAddins.Commands.LayerManagerCommand))]

namespace SKAcadAddins.Commands
{
    public class LayerManagerCommand
    {
        [CommandMethod("LX")]
        public void ShowForm()
        {
            FormsApp.EnableVisualStyles();
            LayerManagerForm form = new LayerManagerForm();
            AcadApp.ShowModalDialog(form);
        }
    }

    public class LayerManagerForm : Form
    {
        private TextBox txtFilePath;
        private Button btnBrowse, btnExport, btnImport;
        private Label lblStatus;
        static LayerManagerForm()
        {
            // Không cần thiết lập license cho NPOI
        }

        public LayerManagerForm()
        {
            Text = "Shinken Group® - Layer Manager";
            Size = new Size(440, 180);
            StartPosition = FormStartPosition.CenterScreen;

            Label lbl = new Label
            {
                Text = "Excel File Path:",
                Location = new Point(10, 20),
                AutoSize = true
            };

            txtFilePath = new TextBox
            {
                Location = new Point(100, 15),
                Width = 250
            };

            btnBrowse = new Button
            {
                Text = "Browse",
                Location = new Point(360, 13),
                Width = 60
            };
            btnBrowse.Click += BtnBrowse_Click;

            btnExport = new Button
            {
                Text = "Export Layers",
                Location = new Point(30, 60),
                Size = new Size(120, 35)
            };
            btnExport.Click += BtnExport_Click;

            btnImport = new Button
            {
                Text = "Import Layers",
                Location = new Point(170, 60),
                Size = new Size(120, 35)
            };
            btnImport.Click += BtnImport_Click;

            lblStatus = new Label
            {
                Text = "",
                ForeColor = AutoColor.DarkGreen,
                Location = new Point(30, 105),
                AutoSize = true
            };

            Controls.Add(lbl);
            Controls.Add(txtFilePath);
            Controls.Add(btnBrowse);
            Controls.Add(btnExport);
            Controls.Add(btnImport);
            Controls.Add(lblStatus);
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Select Excel File"
            };
            if (dlg.ShowDialog() == DialogResult.OK)
                txtFilePath.Text = dlg.FileName;
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFilePath.Text)) return;

            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                    IWorkbook workbook = new XSSFWorkbook();
                    ISheet sheet1 = workbook.CreateSheet("Layers");
                    ISheet sheet2 = workbook.CreateSheet("LayerRename");

                    string[] headers1 = { "LayerName", "ColorARGB", "Linetype", "IsPlottable", "IsFrozen", "IsOff", "IsLocked" };
                    IRow headerRow1 = sheet1.CreateRow(0);
                    IRow headerRow2 = sheet2.CreateRow(0);
                    headerRow2.CreateCell(0).SetCellValue("LayerName");
                    for (int i = 0; i < headers1.Length; i++)
                    {
                        headerRow1.CreateCell(i).SetCellValue(headers1[i]);
                    }

                    int row = 1;
                    foreach (ObjectId id in layerTable)
                    {
                        LayerTableRecord layer = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                        IRow dataRow1 = sheet1.CreateRow(row);
                        IRow dataRow2 = sheet2.CreateRow(row);
                        dataRow1.CreateCell(0).SetCellValue(layer.Name);
                        dataRow1.CreateCell(1).SetCellValue(layer.Color.ColorValue.ToArgb());
                        dataRow1.CreateCell(2).SetCellValue(layer.LinetypeObjectId.ToString());
                        dataRow1.CreateCell(3).SetCellValue(layer.IsPlottable);
                        dataRow1.CreateCell(4).SetCellValue(layer.IsFrozen);
                        dataRow1.CreateCell(5).SetCellValue(layer.IsOff);
                        dataRow1.CreateCell(6).SetCellValue(layer.IsLocked);
                        dataRow2.CreateCell(0).SetCellValue(""); // Để trống cho người dùng nhập tên mới nếu muốn
                        row++;
                    }

                    using (var fs = new FileStream(txtFilePath.Text, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(fs);
                    }

                    tr.Commit();
                }

                lblStatus.ForeColor = AutoColor.Green;
                lblStatus.Text = "✅ Xuất layer thành công!";
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = AutoColor.Red;
                lblStatus.Text = "❌ " + ex.Message;
            }
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFilePath.Text) || !File.Exists(txtFilePath.Text))
            {
                lblStatus.ForeColor = AutoColor.Red;
                lblStatus.Text = "❌ File không tồn tại.";
                return;
            }

            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                    using (var fs = new FileStream(txtFilePath.Text, FileMode.Open, FileAccess.Read))
                    {
                        IWorkbook workbook = new XSSFWorkbook(fs);
                        ISheet sheetRename = workbook.GetSheet("LayerRename");
                        ISheet sheetLayers = workbook.GetSheet("Layers");
                        if (sheetRename == null || sheetLayers == null)
                        {
                            lblStatus.ForeColor = AutoColor.Red;
                            lblStatus.Text = "❌ Không tìm thấy sheet phù hợp.";
                            return;
                        }
                        int rowCount = sheetLayers.LastRowNum;
                        for (int row = 1; row <= rowCount; row++)
                        {
                            IRow dataRow = sheetLayers.GetRow(row);
                            IRow renameRow = sheetRename.GetRow(row);
                            if (dataRow == null) continue;
                            string name = dataRow.GetCell(0)?.ToString();
                            if (string.IsNullOrEmpty(name) || !layerTable.Has(name)) continue;
                            LayerTableRecord layer = (LayerTableRecord)tr.GetObject(layerTable[name], OpenMode.ForWrite);
                            string newName = renameRow?.GetCell(0)?.ToString();
                            if (!string.IsNullOrEmpty(newName) && newName != name && !layerTable.Has(newName))
                            {
                                layer.Name = newName;
                            }
                            if (int.TryParse(dataRow.GetCell(1)?.ToString(), out int argb))
                                layer.Color = Color.FromColor(AutoColor.FromArgb(argb));
                            string linetype = dataRow.GetCell(2)?.ToString();
                            // TODO: Nếu cần set lại linetype thì bổ sung logic
                            layer.IsPlottable = bool.TryParse(dataRow.GetCell(3)?.ToString(), out bool p) && p;
                            layer.IsFrozen = bool.TryParse(dataRow.GetCell(4)?.ToString(), out bool f) && f;
                            layer.IsOff = bool.TryParse(dataRow.GetCell(5)?.ToString(), out bool o) && o;
                            layer.IsLocked = bool.TryParse(dataRow.GetCell(6)?.ToString(), out bool l) && l;
                        }
                    }
                    tr.Commit();
                }

                lblStatus.ForeColor = AutoColor.Green;
                lblStatus.Text = "✅ Cập nhật layer thành công!";
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = AutoColor.Red;
                lblStatus.Text = "❌ " + ex.Message;
            }
        }
    }
}
