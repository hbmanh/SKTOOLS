using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using Microsoft.Win32;
using System.IO;
using SKRevitAddins.ViewModel;

namespace SKRevitAddins.Commands.ExportSchedulesToExcel
{
    public class ExportSchedulesToExcelRequestHandler : IExternalEventHandler
    {
        private ExportSchedulesToExcelViewModel _vm;
        private ExportSchedulesToExcelRequest _request = new ExportSchedulesToExcelRequest();

        public ExportSchedulesToExcelRequestHandler(ExportSchedulesToExcelViewModel viewModel)
        {
            _vm = viewModel;
        }

        public ExportSchedulesToExcelRequest Request => _request;

        public void Execute(UIApplication uiApp)
        {
            try
            {
                Document doc = uiApp.ActiveUIDocument.Document;
                var reqId = _request.Take();
                switch (reqId)
                {
                    case RequestId.None:
                        break;
                    case RequestId.Export:
                        DoExport(doc);
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export Error", ex.Message);
            }
        }

        public string GetName()
        {
            return "ExportSchedulesToExcelRequestHandler";
        }

        private void DoExport(Document doc)
        {
            if (_vm.SelectedSchedules == null || !_vm.SelectedSchedules.Any())
            {
                TaskDialog.Show("Export Schedules", "No schedules selected.");
                return;
            }

            // Hỏi người dùng lưu file Excel ở đâu (sử dụng SaveFileDialog của WPF)
            var saveDialog = new SaveFileDialog()
            {
                Title = "Save Excel File",
                Filter = "Excel File (*.xlsx)|*.xlsx",
                FileName = "SelectedSchedules.xlsx",
                DefaultExt = ".xlsx"
            };

            bool? result = saveDialog.ShowDialog();
            if (result != true)
            {
                return;
            }

            string excelFilePath = saveDialog.FileName;

            // Xuất các schedule được chọn vào file Excel
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage())
            {
                foreach (var scheduleItem in _vm.SelectedSchedules)
                {
                    ViewSchedule schedule = scheduleItem.Schedule;
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

            TaskDialog.Show("Export Schedules", $"Export completed!\nExcel file: {excelFilePath}");
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
}
