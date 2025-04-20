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
        private Button btnExport, btnImport;
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
            txtFilePath.ReadOnly = true;

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
            Controls.Add(btnExport);
            Controls.Add(btnImport);
            Controls.Add(lblStatus);
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Chọn nơi lưu file Excel"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            txtFilePath.Text = dlg.FileName;

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
                    for (int i = 0; i < headers1.Length; i++)
                    {
                        headerRow1.CreateCell(i).SetCellValue(headers1[i]);
                        headerRow2.CreateCell(i).SetCellValue(headers1[i]);
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
                        dataRow2.CreateCell(0).SetCellValue(layer.Name);
                        dataRow2.CreateCell(1).SetCellValue(layer.Color.ColorValue.ToArgb());
                        dataRow2.CreateCell(2).SetCellValue(layer.LinetypeObjectId.ToString());
                        dataRow2.CreateCell(3).SetCellValue(layer.IsPlottable);
                        dataRow2.CreateCell(4).SetCellValue(layer.IsFrozen);
                        dataRow2.CreateCell(5).SetCellValue(layer.IsOff);
                        dataRow2.CreateCell(6).SetCellValue(layer.IsLocked);
                        row++;
                    }

                    using (var fs = new FileStream(txtFilePath.Text, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
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
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Chọn file Excel để nhập"
            };
            if (dlg.ShowDialog() != DialogResult.OK)
            {
                lblStatus.ForeColor = AutoColor.Red;
                lblStatus.Text = "❌ Chưa chọn file để nhập.";
                return;
            }
            txtFilePath.Text = dlg.FileName;
            if (!File.Exists(txtFilePath.Text))
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
                    using (var fs = new FileStream(txtFilePath.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
                        int rowCount1 = sheetLayers.LastRowNum;
                        int rowCount2 = sheetRename.LastRowNum;
                        if (rowCount1 != rowCount2)
                        {
                            lblStatus.ForeColor = AutoColor.Red;
                            lblStatus.Text = "❌ Số dòng giữa hai sheet không khớp.";
                            return;
                        }
                        int colCount1 = sheetLayers.GetRow(0)?.LastCellNum ?? 0;
                        int colCount2 = sheetRename.GetRow(0)?.LastCellNum ?? 0;
                        if (colCount1 != colCount2 || colCount1 < 7)
                        {
                            lblStatus.ForeColor = AutoColor.Red;
                            lblStatus.Text = "❌ Số cột không hợp lệ hoặc thiếu dữ liệu.";
                            return;
                        }
                        for (int row = 1; row <= rowCount1; row++)
                        {
                            IRow dataRow1 = sheetLayers.GetRow(row);
                            IRow dataRow2 = sheetRename.GetRow(row);
                            if (dataRow1 == null || dataRow2 == null)
                            {
                                lblStatus.ForeColor = AutoColor.Red;
                                lblStatus.Text = $"❌ Dòng {row + 1} bị thiếu dữ liệu.";
                                return;
                            }
                            string name = dataRow1.GetCell(0)?.ToString();
                            if (string.IsNullOrEmpty(name) || !layerTable.Has(name)) continue;
                            LayerTableRecord layer = (LayerTableRecord)tr.GetObject(layerTable[name], OpenMode.ForWrite);
                            // Name
                            string newName = dataRow2.GetCell(0)?.ToString();
                            if (!string.IsNullOrEmpty(newName) && !newName.Equals(name, StringComparison.OrdinalIgnoreCase))
                            {
                                if (name.Equals("Layer1", StringComparison.OrdinalIgnoreCase))
                                    continue;
                                char[] invalidChars = { '\\', '/', ':', ';', '<', '>', '?', '"', '|', '=', '`', '*', ',', ' ', '\t' };
                                if (newName.IndexOfAny(invalidChars) >= 0)
                                {
                                    lblStatus.ForeColor = AutoColor.Red;
                                    lblStatus.Text = $"❌ Tên layer mới ở dòng {row + 1} chứa ký tự không hợp lệ: '{newName}'. Không được chứa các ký tự: \\ / : ; < > ? \" | = ` * , hoặc khoảng trắng.";
                                    return;
                                }
                                if (string.IsNullOrWhiteSpace(newName))
                                {
                                    lblStatus.ForeColor = AutoColor.Red;
                                    lblStatus.Text = $"❌ Tên layer mới ở dòng {row + 1} bị bỏ trống.";
                                    return;
                                }
                                bool nameExists = false;
                                foreach (ObjectId lid in layerTable)
                                {
                                    LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(lid, OpenMode.ForRead);
                                    if (ltr.Name.Equals(newName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        nameExists = true;
                                        break;
                                    }
                                }
                                if (nameExists)
                                {
                                    lblStatus.ForeColor = AutoColor.Red;
                                    lblStatus.Text = $"❌ Tên layer mới ở dòng {row + 1} đã tồn tại: '{newName}'. Vui lòng chọn tên khác.";
                                    return;
                                }
                                try
                                {
                                    layer.Name = newName;
                                }
                                catch (System.Exception ex)
                                {
                                    lblStatus.ForeColor = AutoColor.Red;
                                    lblStatus.Text = $"❌ Đổi tên layer ở dòng {row + 1} thất bại cho tên '{newName}': {ex.Message}. Hãy kiểm tra lại tên layer trong file Excel.";
                                    return;
                                }
                            }
                            // Color
                            if (int.TryParse(dataRow1.GetCell(1)?.ToString(), out int argb1) && int.TryParse(dataRow2.GetCell(1)?.ToString(), out int argb2))
                            {
                                if (argb1 != argb2)
                                    layer.Color = Color.FromColor(AutoColor.FromArgb(argb2));
                            }
                            // Linetype
                            string linetype1 = dataRow1.GetCell(2)?.ToString();
                            string linetype2 = dataRow2.GetCell(2)?.ToString();
                            if (!string.IsNullOrEmpty(linetype2) && linetype1 != linetype2)
                            {
                                // TODO: Nếu cần set lại linetype thì bổ sung logic
                            }
                            // IsPlottable
                            bool.TryParse(dataRow1.GetCell(3)?.ToString(), out bool p1);
                            bool.TryParse(dataRow2.GetCell(3)?.ToString(), out bool p2);
                            if (p1 != p2) layer.IsPlottable = p2;
                            // IsFrozen
                            bool.TryParse(dataRow1.GetCell(4)?.ToString(), out bool f1);
                            bool.TryParse(dataRow2.GetCell(4)?.ToString(), out bool f2);
                            if (f1 != f2) layer.IsFrozen = f2;
                            // IsOff
                            bool.TryParse(dataRow1.GetCell(5)?.ToString(), out bool o1);
                            bool.TryParse(dataRow2.GetCell(5)?.ToString(), out bool o2);
                            if (o1 != o2) layer.IsOff = o2;
                            // IsLocked
                            bool.TryParse(dataRow1.GetCell(6)?.ToString(), out bool l1);
                            bool.TryParse(dataRow2.GetCell(6)?.ToString(), out bool l2);
                            if (l1 != l2) layer.IsLocked = l2;
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
