using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using SKRevitAddins.ViewModel;
using SKRevitAddins.ExportSchedulesToExcel;

namespace SKRevitAddins.ExportSchedulesToExcel
{
    public class ExportSchedulesToExcelRequestHandler : IExternalEventHandler
    {
        private ExportSchedulesToExcelRequest _request;
        private ExportSchedulesToExcelViewModel _vm;
        private CancellationTokenSource _cts;
        // Sử dụng Dictionary để tối ưu tên sheet duy nhất
        private Dictionary<string, int> sheetNameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        public ExportSchedulesToExcelRequestHandler(ExportSchedulesToExcelViewModel viewModel, ExportSchedulesToExcelRequest request)
        {
            _vm = viewModel;
            _request = request;
            _cts = new CancellationTokenSource();
        }

        // Cho phép code-behind gọi: handler.Request.Make(RequestId.Export)
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
                TaskDialog.Show("Export Error", ex.ToString());
                LogException(ex);
            }
        }

        private void DoExport(UIApplication uiApp)
        {
            _vm.ExportStatusMessage = "";
            if (_vm.SelectedSchedules == null || _vm.SelectedSchedules.Count == 0)
            {
                _vm.ExportStatusMessage = "No schedule selected to export.";
                return;
            }

            var sfd = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save Excel File",
                Filter = "Excel File (*.xlsx)|*.xlsx",
                FileName = "SelectedSchedules.xlsx"
            };

            bool? result = sfd.ShowDialog();
            if (result != true)
            {
                _vm.ExportStatusMessage = "Export cancelled by user.";
                return;
            }

            string excelFilePath = sfd.FileName;
            // Reset lại từ điển tên sheet
            sheetNameCounts.Clear();

            // Thu thập dữ liệu từ các schedule trên luồng chính (vì cần gọi API Revit)
            List<(string ScheduleName, List<List<string>> Data)> scheduleDataList = new List<(string, List<List<string>>)>();
            foreach (var schedItem in _vm.SelectedSchedules)
            {
                var data = GetScheduleData(schedItem.Schedule);
                if (data.Count > 0)
                {
                    scheduleDataList.Add((schedItem.Name, data));
                }
            }

            try
            {
                // Xử lý xuất Excel bất đồng bộ để không block giao diện
                Task.Run(() =>
                {
                    CancellationToken token = _cts.Token;
                    OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (var package = new OfficeOpenXml.ExcelPackage())
                    {
                        int total = scheduleDataList.Count;
                        for (int i = 0; i < total; i++)
                        {
                            token.ThrowIfCancellationRequested();
                            var item = scheduleDataList[i];
                            string baseName = CleanSheetName(item.ScheduleName);
                            string sheetName = GetUniqueSheetName(baseName);
                            var ws = package.Workbook.Worksheets.Add(sheetName);
                            var data = item.Data;
                            for (int r = 0; r < data.Count; r++)
                            {
                                token.ThrowIfCancellationRequested();
                                for (int c = 0; c < data[r].Count; c++)
                                {
                                    ws.Cells[r + 1, c + 1].Value = data[r][c];
                                }
                            }

                            // Tạo table nếu có header
                            if (data[0].Count > 0)
                            {
                                var range = ws.Cells[1, 1, data.Count, data[0].Count];
                                string tableName = CleanTableName(item.ScheduleName);
                                try
                                {
                                    var tbl = ws.Tables.Add(range, tableName);
                                    tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Medium6;
                                }
                                catch (Exception ex)
                                {
                                    Log("Error creating table for " + item.ScheduleName + ": " + ex.Message);
                                }
                            }
                            ws.Cells[ws.Dimension.Address].AutoFitColumns();

                            // Cho phép hệ thống "thở" mỗi 10 vòng lặp
                            if (i % 10 == 0)
                            {
                                Task.Delay(1, token).Wait(token);
                            }
                        }
                        FileInfo fi = new FileInfo(excelFilePath);
                        package.SaveAs(fi);
                    }
                }, _cts.Token).GetAwaiter().GetResult();

                _vm.ExportStatusMessage = "Export completed!";
            }
            catch (OperationCanceledException)
            {
                _vm.ExportStatusMessage = "Export cancelled by user.";
                Log("Export cancelled by user.");
            }
            catch (Exception ex)
            {
                _vm.ExportStatusMessage = "Export failed: " + ex.Message;
                LogException(ex);
            }
        }

        // Lấy dữ liệu từ schedule (chạy trên luồng chính)
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

        // Làm sạch tên sheet sử dụng Regex
        private string CleanSheetName(string name)
        {
            if (string.IsNullOrEmpty(name))
                name = "Sheet";
            string invalidChars = new string(Path.GetInvalidFileNameChars());
            string pattern = $"[{Regex.Escape(invalidChars)}]";
            string cleaned = Regex.Replace(name, pattern, "_");
            if (cleaned.Length > 31)
                cleaned = cleaned.Substring(0, 31);
            if (string.IsNullOrWhiteSpace(cleaned))
                cleaned = "Sheet";
            return cleaned;
        }

        // Làm sạch tên table sử dụng Regex
        private string CleanTableName(string name)
        {
            if (string.IsNullOrEmpty(name))
                name = "Table";
            string result = Regex.Replace(name, @"\s+", "_");
            result = new string(result.Where(ch => char.IsLetterOrDigit(ch) || ch == '_').ToArray());
            if (string.IsNullOrEmpty(result) || (!char.IsLetter(result[0]) && result[0] != '_'))
            {
                result = "_" + result;
            }
            return result;
        }

        // Tối ưu GetUniqueSheetName sử dụng Dictionary để theo dõi số lần xuất hiện của baseName
        private string GetUniqueSheetName(string baseName)
        {
            if (!sheetNameCounts.ContainsKey(baseName))
            {
                sheetNameCounts[baseName] = 0;
                return baseName;
            }
            else
            {
                sheetNameCounts[baseName]++;
                string newName = $"{baseName}_{sheetNameCounts[baseName]}";
                if (newName.Length > 31)
                    newName = newName.Substring(0, 31);
                return newName;
            }
        }

        // Các phương thức ghi log đơn giản (không tạo lớp mới)
        private void Log(string message)
        {
            try
            {
                string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExportSchedulesLog.txt");
                File.AppendAllText(logFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}{Environment.NewLine}");
            }
            catch { }
        }

        private void LogException(Exception ex)
        {
            Log("Exception: " + ex.ToString());
        }
    }
}
