using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Linq;
using System.Collections.Generic;
using SKRevitAddins.ViewModel;

namespace SKRevitAddins.Commands.ExportSchedulesToExcel
{
    public class ExportSchedulesToExcelRequestHandler : IExternalEventHandler
    {
        private ExportSchedulesToExcelRequest _request;
        private ExportSchedulesToExcelViewModel _vm;

        public ExportSchedulesToExcelRequestHandler(
            ExportSchedulesToExcelViewModel viewModel,
            ExportSchedulesToExcelRequest request)
        {
            _vm = viewModel;
            _request = request;
        }

        // Cho phép code-behind gọi handler.Request.Make(...)
        public ExportSchedulesToExcelRequest Request => _request;

        public string GetName() => "ExportSchedulesToExcelRequestHandler";

        public void Execute(UIApplication uiApp)
        {
            try
            {
                var reqId = _request.Take();
                switch (reqId)
                {
                    case RequestId.None:
                        break;
                    case RequestId.Export:
                        DoExport(uiApp);
                        break;
                }
            }
            catch (Exception ex)
            {
                // Tuỳ ý ghi log hoặc hiển thị lỗi
            }
        }

        private void DoExport(UIApplication uiApp)
        {
            // Xoá thông báo cũ (nếu có)
            _vm.ExportStatusMessage = "";

            var doc = uiApp.ActiveUIDocument.Document;
            var selected = _vm.SelectedSchedules?.ToList();
            if (selected == null || selected.Count == 0)
            {
                _vm.ExportStatusMessage = "No schedule selected to export.";
                return;
            }

            // Hỏi người dùng nơi lưu file
            var sfd = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save Excel File",
                Filter = "Excel File (*.xlsx)|*.xlsx",
                FileName = "SelectedSchedules.xlsx"
            };
            bool? result = sfd.ShowDialog();
            if (result != true)
            {
                // Người dùng bấm Cancel => không làm gì
                return;
            }

            string excelFilePath = sfd.FileName;

            // Tiến hành Export
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage())
            {
                foreach (var schedItem in selected)
                {
                    var vs = schedItem.Schedule;
                    var data = GetScheduleData(vs);
                    if (data.Count == 0) continue;

                    string sheetName = CleanSheetName(schedItem.Name);
                    var ws = package.Workbook.Worksheets.Add(sheetName);

                    // Ghi dữ liệu
                    for (int r = 0; r < data.Count; r++)
                    {
                        for (int c = 0; c < data[r].Count; c++)
                        {
                            ws.Cells[r + 1, c + 1].Value = data[r][c];
                        }
                    }

                    // Table style
                    if (data[0].Count > 0)
                    {
                        var range = ws.Cells[1, 1, data.Count, data[0].Count];
                        var tbl = ws.Tables.Add(range, CleanTableName(schedItem.Name));
                        tbl.TableStyle = TableStyles.Medium6;
                    }
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                }

                // Lưu file
                var fi = new FileInfo(excelFilePath);
                package.SaveAs(fi);
            }

            // Hiển thị thông báo đã hoàn thành ngay trên cửa sổ chính
            _vm.ExportStatusMessage = "Export completed!";
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
            var invalid = Path.GetInvalidFileNameChars();
            foreach (char c in invalid)
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
            if (string.IsNullOrEmpty(result) || (!char.IsLetter(result[0]) && result[0] != '_'))
            {
                result = "_" + result;
            }
            return result;
        }
    }
}
