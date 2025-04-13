using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using Form = System.Windows.Forms.Form;
using View = Autodesk.Revit.DB.View;
using Point = System.Drawing.Point;
using Color = System.Drawing.Color;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;

namespace SKRevitAddins.Commands.CreateSheetsFromExcel
{
    public class ProgressForm : Form
    {
        public ProgressBar progressBar;

        public ProgressForm(int maxValue)
        {
            Text = "Shinken Group®";
            Size = new Size(420, 100);
            StartPosition = FormStartPosition.CenterScreen;

            PictureBox logo = new PictureBox
            {
                Image = Image.FromFile("C:\\ProgramData\\Autodesk\\Revit\\Addins\\2023\\SKTools.bundle\\Contents\\Resources\\Images\\shinken.png"),
                Size = new Size(32, 32),
                Location = new Point(10, 10),
                SizeMode = PictureBoxSizeMode.StretchImage
            };

            progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Continuous,
                Maximum = maxValue,
                Minimum = 0,
                Value = 0,
                Location = new Point(50, 20),
                Size = new Size(340, 25),
                ForeColor = Color.SkyBlue
            };

            Controls.Add(logo);
            Controls.Add(progressBar);
        }

        public void UpdateProgress(int value)
        {
            progressBar.Value = value;
            Refresh();
        }
    }

    public class SheetSelectionForm : Form
    {
        public List<string> SelectedSheets = new List<string>();
        private DataGridView sheetGrid;

        public SheetSelectionForm(List<(string number, string name)> existingSheets)
        {
            Text = "Shinken Group® - Chọn sheet để tạo lại";
            Size = new Size(500, 500);
            StartPosition = FormStartPosition.CenterScreen;

            PictureBox logo = new PictureBox
            {
                Image = Image.FromFile("C:\\ProgramData\\Autodesk\\Revit\\Addins\\2023\\SKTools.bundle\\Contents\\Resources\\Images\\shinken.png"),
                Size = new Size(32, 32),
                Location = new Point(10, 10),
                SizeMode = PictureBoxSizeMode.StretchImage
            };

            Label label = new Label
            {
                Text = "Các sheet đã tồn tại, chọn để tạo lại:",
                Location = new Point(50, 15),
                AutoSize = true
            };

            sheetGrid = new DataGridView
            {
                Location = new Point(20, 50),
                Size = new Size(440, 320),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                MultiSelect = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false
            };

            sheetGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Sheet Number", Width = 150 });
            sheetGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Sheet Name", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });

            foreach (var sheet in existingSheets)
                sheetGrid.Rows.Add(sheet.number, sheet.name);

            Button btnOk = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new Point(200, 390),
                Width = 75
            };

            Button btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(300, 390),
                Width = 75
            };

            Controls.Add(logo);
            Controls.Add(label);
            Controls.Add(sheetGrid);
            Controls.Add(btnOk);
            Controls.Add(btnCancel);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SelectedSheets.Clear();
            foreach (DataGridViewRow row in sheetGrid.SelectedRows)
            {
                SelectedSheets.Add(row.Cells[0].Value.ToString());
            }
            base.OnFormClosing(e);
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class CreateSheetsFromExcelCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Application.EnableVisualStyles();

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel files (*.xlsx)|*.xlsx" };
            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;

            string excelPath = openFileDialog.FileName;
            FileInfo fileInfo = new FileInfo(excelPath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FamilySymbol titleBlockType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilySymbol>()
                .FirstOrDefault(fs => fs.Family.Name == "AWS-Titleblock-EMEA-APAC-Horizontal_NRT085_250403" && fs.Name == "ISO A0 - 841mm x 1189mm");

            if (titleBlockType == null)
            {
                message = "Không tìm thấy Title Block đã chỉ định.";
                return Result.Failed;
            }

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                List<(string, string)> duplicateSheets = new List<(string, string)>();
                Dictionary<int, (string, string, string)> sheetData = new Dictionary<int, (string, string, string)>();

                for (int row = 2; row <= rowCount; row++)
                {
                    string sheetNumber = worksheet.Cells[row, 1].Text.Trim();
                    string sheetName = worksheet.Cells[row, 2].Text.Trim();
                    string viewGroup = worksheet.Cells[row, 3].Text.Trim();
                    if (string.IsNullOrWhiteSpace(sheetNumber) || string.IsNullOrWhiteSpace(sheetName)) continue;
                    if (new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>().Any(s => s.SheetNumber == sheetNumber))
                        duplicateSheets.Add((sheetNumber, sheetName));
                    else
                        sheetData[row] = (sheetNumber, sheetName, viewGroup);
                }

                if (duplicateSheets.Any())
                {
                    var selectionForm = new SheetSelectionForm(duplicateSheets);
                    if (selectionForm.ShowDialog() == DialogResult.OK)
                    {
                        var sheetsToDelete = selectionForm.SelectedSheets;
                        using (Transaction delTrans = new Transaction(doc, "Xóa các sheet cũ"))
                        {
                            delTrans.Start();
                            foreach (string sheetNum in sheetsToDelete)
                            {
                                var existingSheet = new FilteredElementCollector(doc)
                                    .OfClass(typeof(ViewSheet)).Cast<ViewSheet>()
                                    .FirstOrDefault(s => s.SheetNumber == sheetNum);
                                if (existingSheet != null) doc.Delete(existingSheet.Id);
                            }
                            delTrans.Commit();
                        }
                        for (int row = 2; row <= rowCount; row++)
                        {
                            string sheetNumber = worksheet.Cells[row, 1].Text.Trim();
                            string sheetName = worksheet.Cells[row, 2].Text.Trim();
                            string viewGroup = worksheet.Cells[row, 3].Text.Trim();
                            if (sheetsToDelete.Contains(sheetNumber))
                                sheetData[row] = (sheetNumber, sheetName, viewGroup);
                        }
                    }
                }

                ProgressForm progressForm = new ProgressForm(sheetData.Count);
                progressForm.Show();
                progressForm.Refresh();

                using (Transaction trans = new Transaction(doc, "Create Sheets and Place Views"))
                {
                    trans.Start();
                    int index = 0;
                    foreach (var entry in sheetData)
                    {
                        var (sheetNumber, sheetName, viewGroup) = entry.Value;
                        ViewSheet newSheet = ViewSheet.Create(doc, titleBlockType.Id);
                        newSheet.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set(sheetNumber);
                        newSheet.get_Parameter(BuiltInParameter.SHEET_NAME).Set(sheetName);
                        var param = newSheet.LookupParameter("Sub-Package");
                        if (param != null && !param.IsReadOnly) param.Set(viewGroup);
                        string viewFullName = sheetNumber + " - " + sheetName;
                        var viewToPlace = new FilteredElementCollector(doc)
                            .OfClass(typeof(View)).Cast<View>()
                            .FirstOrDefault(v => v.Name.Equals(viewFullName));
                        if (viewToPlace != null && Viewport.CanAddViewToSheet(doc, newSheet.Id, viewToPlace.Id))
                        {
                            var sheetBBox = newSheet.Outline;
                            var sheetCenter = new XYZ((sheetBBox.Min.U - sheetBBox.Max.U) / 2,
                                                      (sheetBBox.Min.V + sheetBBox.Max.V) / 2, 0);
                            Viewport.Create(doc, newSheet.Id, viewToPlace.Id, sheetCenter);
                        }
                        progressForm.UpdateProgress(++index);
                    }
                    trans.Commit();
                }
                progressForm.Close();
            }
            TaskDialog.Show("Thông báo", "Đã tạo Sheet và đặt View hoàn tất!");
            return Result.Succeeded;
        }
    }
}
