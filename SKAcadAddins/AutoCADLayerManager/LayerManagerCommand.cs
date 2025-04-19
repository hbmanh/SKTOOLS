using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Colors;
using OfficeOpenXml;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

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
        [CommandMethod("LayerManagerUI")]
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

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

                    using (ExcelPackage package = new ExcelPackage())
                    {
                        var sheet1 = package.Workbook.Worksheets.Add("Layers");
                        var sheet2 = package.Workbook.Worksheets.Add("Metadata");

                        string[] headers = { "LayerName", "ColorARGB", "Linetype", "IsPlottable", "IsFrozen", "IsOff", "IsLocked" };
                        for (int i = 0; i < headers.Length; i++)
                            sheet1.Cells[1, i + 1].Value = headers[i];

                        int row = 2;
                        foreach (ObjectId id in layerTable)
                        {
                            LayerTableRecord layer = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                            sheet1.Cells[row, 1].Value = layer.Name;
                            sheet1.Cells[row, 2].Value = layer.Color.ColorValue.ToArgb();
                            sheet1.Cells[row, 3].Value = layer.LinetypeObjectId.ToString();
                            sheet1.Cells[row, 4].Value = layer.IsPlottable;
                            sheet1.Cells[row, 5].Value = layer.IsFrozen;
                            sheet1.Cells[row, 6].Value = layer.IsOff;
                            sheet1.Cells[row, 7].Value = layer.IsLocked;
                            row++;
                        }

                        sheet2.Cells[1, 1].Value = "ColorARGB: int ARGB value";
                        sheet2.Cells[2, 1].Value = "Linetype: BYLAYER, CONTINUOUS...";
                        sheet2.Cells[3, 1].Value = "Boolean fields: True/False";

                        package.SaveAs(new FileInfo(txtFilePath.Text));
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
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(txtFilePath.Text)))
                    {
                        var sheet = package.Workbook.Worksheets["Layers"];
                        int rowCount = sheet.Dimension.End.Row;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            string name = sheet.Cells[row, 1].Text;
                            if (!layerTable.Has(name)) continue;

                            LayerTableRecord layer = (LayerTableRecord)tr.GetObject(layerTable[name], OpenMode.ForWrite);

                            if (int.TryParse(sheet.Cells[row, 2].Text, out int argb))
                                layer.Color = Color.FromColor(AutoColor.FromArgb(argb)); // ✅ đúng AutoCAD.Colors.Color

                            string linetype = sheet.Cells[row, 3].Text;
                            if (!string.IsNullOrEmpty(linetype))
                                layer.LinetypeObjectId.ToString(); 

                            layer.IsPlottable = bool.TryParse(sheet.Cells[row, 4].Text, out bool p) && p;
                            layer.IsFrozen = bool.TryParse(sheet.Cells[row, 5].Text, out bool f) && f;
                            layer.IsOff = bool.TryParse(sheet.Cells[row, 6].Text, out bool o) && o;
                            layer.IsLocked = bool.TryParse(sheet.Cells[row, 7].Text, out bool l) && l;
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
