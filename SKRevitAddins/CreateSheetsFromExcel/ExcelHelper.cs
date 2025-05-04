using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using HorizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment;

namespace SKRevitAddins.CreateSheetsFromExcel
{
    public static class ExcelHelper
    {
        public static void CreateTemplateExcel(string filePath)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheets");

            IRow header = sheet.CreateRow(0);
            string[] headers = {
                "Sheet Number",
                "Sheet Name",
                "View Group",
                "Create View",  // FL, SL, DV
                "Level"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = header.CreateCell(i);
                cell.SetCellValue(headers[i]);

                var style = workbook.CreateCellStyle();
                var font = workbook.CreateFont();
                font.IsBold = true;
                style.SetFont(font);
                style.FillForegroundColor = IndexedColors.LightOrange.Index;
                style.FillPattern = FillPattern.SolidForeground;
                style.Alignment = HorizontalAlignment.Center;
                style.VerticalAlignment = VerticalAlignment.Center;

                cell.CellStyle = style;
                sheet.AutoSizeColumn(i);
            }

            try
            {
                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    workbook.Write(fs);
                }

                // Mở file Excel trực tiếp
                Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể tạo/mở file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public static Dictionary<int, (string number, string name, string group, string createView, string level)> ReadExcel(string filePath)
        {
            var data = new Dictionary<int, (string, string, string, string, string)>();

            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook wb = new XSSFWorkbook(fs);
                    ISheet sheet = wb.GetSheetAt(0);
                    int rowCount = sheet.LastRowNum;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        var row = sheet.GetRow(i);
                        if (row == null) continue;

                        string number = row.GetCell(0)?.ToString().Trim();
                        string name = row.GetCell(1)?.ToString().Trim();
                        string group = row.GetCell(2)?.ToString().Trim();
                        string createView = row.GetCell(3)?.ToString().Trim().ToUpper();
                        string level = row.GetCell(4)?.ToString().Trim();

                        if (string.IsNullOrWhiteSpace(number) || string.IsNullOrWhiteSpace(name))
                            continue;

                        data[i] = (number, name, group, createView, level);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi đọc Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return data;
        }
    }
}
