using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using OfficeOpenXml.Table;
// WinForm
// Để chỉnh kích cỡ, Dock, etc.
using Form = System.Windows.Forms.Form;
using Panel = System.Windows.Forms.Panel;

namespace SKRevitAddins
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // (1) Thu thập tất cả ViewSchedule (bỏ qua template)
                List<ViewSchedule> allSchedules = GetAllSchedules(doc);
                if (!allSchedules.Any())
                {
                    TaskDialog.Show("Export Schedules", "No schedules found in this project.");
                    return Result.Cancelled;
                }

                // (2) Hiển thị form WinForm để người dùng chọn schedule xuất
                List<ViewSchedule> selectedSchedules;
                using (ScheduleSelectionForm form = new ScheduleSelectionForm(allSchedules))
                {
                    var dr = form.ShowDialog();
                    if (dr != DialogResult.OK)
                    {
                        // Người dùng nhấn Cancel hoặc đóng form => không xuất
                        return Result.Cancelled;
                    }
                    selectedSchedules = form.SelectedSchedules;
                }

                if (!selectedSchedules.Any())
                {
                    TaskDialog.Show("Export Schedules", "No schedules selected.");
                    return Result.Cancelled;
                }

                // (3) Hỏi người dùng lưu file Excel ở đâu (SaveFileDialog)
                string excelFilePath;
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Title = "Save Excel File";
                    sfd.Filter = "Excel File (*.xlsx)|*.xlsx";
                    sfd.FileName = "SelectedSchedules.xlsx"; // Tên gợi ý
                    if (sfd.ShowDialog() != DialogResult.OK)
                    {
                        return Result.Cancelled;
                    }
                    excelFilePath = sfd.FileName;
                }

                // (4) Xuất các schedule đã chọn vào file Excel
                ExportSchedulesToExcel(selectedSchedules, excelFilePath);

                // (5) Thông báo
                TaskDialog.Show("Export Schedules",
                    $"Export completed!\nExcel file: {excelFilePath}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private List<ViewSchedule> GetAllSchedules(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            return collector
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => !vs.IsTemplate)
                .ToList();
        }

        private List<List<string>> GetScheduleData(ViewSchedule schedule)
        {
            var tableData = schedule.GetTableData();
            var section = tableData.GetSectionData(SectionType.Body);

            List<List<string>> rowsData = new List<List<string>>();
            int rowCount = section.NumberOfRows;
            int colCount = section.NumberOfColumns;

            for (int r = 0; r < rowCount; r++)
            {
                List<string> row = new List<string>();
                for (int c = 0; c < colCount; c++)
                {
                    string cellText = schedule.GetCellText(SectionType.Body, r, c);
                    row.Add(cellText);
                }
                rowsData.Add(row);
            }
            return rowsData;
        }

        private void ExportSchedulesToExcel(List<ViewSchedule> schedules, string excelFilePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage())
            {
                foreach (var schedule in schedules)
                {
                    var data = GetScheduleData(schedule);
                    if (data.Count == 0) continue;

                    string wsName = CleanSheetName(schedule.Name);
                    var ws = package.Workbook.Worksheets.Add(wsName);

                    for (int r = 0; r < data.Count; r++)
                    {
                        for (int c = 0; c < data[r].Count; c++)
                        {
                            ws.Cells[r + 1, c + 1].Value = data[r][c];
                        }
                    }

                    if (data[0].Count > 0)
                    {
                        var range = ws.Cells[1, 1, data.Count, data[0].Count];
                        string tableName = CleanTableName(schedule.Name);
                        var tbl = ws.Tables.Add(range, tableName);
                        tbl.TableStyle = TableStyles.Medium6;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                }

                FileInfo fi = new FileInfo(excelFilePath);
                package.SaveAs(fi);
            }
        }

        // Hàm làm sạch tên Worksheet (không liên quan trực tiếp đến lỗi này)
        private string CleanSheetName(string name)
        {
            string invalidChars = new string(Path.GetInvalidFileNameChars()) + @":\/?*[]";
            foreach (char c in invalidChars)
            {
                name = name.Replace(c.ToString(), "_");
            }
            if (name.Length > 31) name = name.Substring(0, 31);
            return name;
        }

        // Cập nhật hàm CleanTableName để đảm bảo tên bắt đầu bằng chữ cái hoặc '_'
        private string CleanTableName(string name)
        {
            string result = name.Replace(" ", "_");
            result = new string(result.Where(ch => char.IsLetterOrDigit(ch) || ch == '_').ToArray());
            if (string.IsNullOrEmpty(result) || (!char.IsLetter(result, 0) && result[0] != '_'))
            {
                result = "_" + result;
            }
            return result;
        }
    }

    public class ScheduleSelectionForm : Form
    {
        private CheckedListBox clbSchedules = new CheckedListBox();
        private Button btnSelectAll = new Button();
        private Button btnDeselectAll = new Button();
        private Button btnOk = new Button();
        private Button btnCancel = new Button();

        private List<ViewSchedule> allSchedules;
        private List<ViewSchedule> selectedSchedules = new List<ViewSchedule>();

        public List<ViewSchedule> SelectedSchedules => selectedSchedules;

        public ScheduleSelectionForm(List<ViewSchedule> schedules)
        {
            this.allSchedules = schedules;
            this.Text = "Select Schedules to Export";
            this.Size = new Size(450, 360);
            this.StartPosition = FormStartPosition.CenterScreen;

            clbSchedules.Dock = DockStyle.Top;
            clbSchedules.Height = 220;
            foreach (var vs in schedules)
            {
                clbSchedules.Items.Add(vs.Name, true);
            }
            this.Controls.Add(clbSchedules);

            Panel panelButtons = new Panel();
            panelButtons.Dock = DockStyle.Bottom;
            panelButtons.Height = 40;
            this.Controls.Add(panelButtons);

            btnSelectAll.Text = "Select All";
            btnSelectAll.Width = 80;
            btnSelectAll.Left = 10;
            btnSelectAll.Top = 5;
            btnSelectAll.Click += (s, e) =>
            {
                for (int i = 0; i < clbSchedules.Items.Count; i++)
                {
                    clbSchedules.SetItemChecked(i, true);
                }
            };
            panelButtons.Controls.Add(btnSelectAll);

            btnDeselectAll.Text = "Deselect All";
            btnDeselectAll.Width = 80;
            btnDeselectAll.Left = 100;
            btnDeselectAll.Top = 5;
            btnDeselectAll.Click += (s, e) =>
            {
                for (int i = 0; i < clbSchedules.Items.Count; i++)
                {
                    clbSchedules.SetItemChecked(i, false);
                }
            };
            panelButtons.Controls.Add(btnDeselectAll);

            btnOk.Text = "OK";
            btnOk.Width = 80;
            btnOk.Left = 200;
            btnOk.Top = 5;
            btnOk.Click += (s, e) =>
            {
                selectedSchedules.Clear();
                for (int i = 0; i < clbSchedules.Items.Count; i++)
                {
                    if (clbSchedules.GetItemChecked(i))
                    {
                        selectedSchedules.Add(allSchedules[i]);
                    }
                }
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
            panelButtons.Controls.Add(btnOk);

            btnCancel.Text = "Cancel";
            btnCancel.Width = 80;
            btnCancel.Left = 290;
            btnCancel.Top = 5;
            btnCancel.Click += (s, e) =>
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };
            panelButtons.Controls.Add(btnCancel);
        }
    }
}
