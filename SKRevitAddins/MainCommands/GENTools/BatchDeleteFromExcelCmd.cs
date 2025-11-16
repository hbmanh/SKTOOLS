using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace SKRevitAddins.GENTools
{
    [Transaction(TransactionMode.Manual)]
    public class BatchDeleteFromExcelCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            string excelPath = PickExcelFile();
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                TaskDialog.Show("Error", "No Excel file selected.");
                return Result.Cancelled;
            }

            int colIndex = PromptColumnChoice();
            if (colIndex < 0)
            {
                TaskDialog.Show("Error", "No column selected.");
                return Result.Cancelled;
            }

            List<int> ids = ReadIdsFromExcel(excelPath, colIndex);

            if (ids.Count == 0)
            {
                TaskDialog.Show("Info", "No valid Element IDs found.");
                return Result.Cancelled;
            }

            int deleted = 0;
            int missing = 0;

            using (Transaction t = new Transaction(doc, "Delete Elements From Excel"))
            {
                t.Start();
                foreach (int id in ids)
                {
                    ElementId eid = new ElementId(id);
                    Element e = doc.GetElement(eid);

                    if (e != null)
                    {
                        doc.Delete(eid);
                        deleted++;
                    }
                    else
                    {
                        missing++;
                    }
                }
                t.Commit();
            }

            TaskDialog.Show("Result",
                $"Total IDs Read: {ids.Count}\n" +
                $"Deleted: {deleted}\n" +
                $"Not Found in Model: {missing}");

            return Result.Succeeded;
        }

        private string PickExcelFile()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls";

            bool? result = dlg.ShowDialog();
            if (result == true)
                return dlg.FileName;

            return null;
        }

        private int PromptColumnChoice()
        {
            TaskDialog td = new TaskDialog("Select Column");
            td.MainInstruction = "Choose Excel Column for Element IDs:";
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "M");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Q");
            td.CommonButtons = TaskDialogCommonButtons.Cancel;

            TaskDialogResult res = td.Show();

            if (res == TaskDialogResult.CommandLink1) return 12;
            if (res == TaskDialogResult.CommandLink2) return 16;

            return -1;
        }

        private List<int> ReadIdsFromExcel(string filePath, int columnIndex)
        {
            List<int> ids = new List<int>();
            IWorkbook workbook;

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                if (filePath.ToLower().EndsWith(".xlsx"))
                    workbook = new XSSFWorkbook(fs);
                else
                    workbook = new HSSFWorkbook(fs);
            }

            ISheet sheet = workbook.GetSheetAt(0);

            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;

                ICell cell = row.GetCell(columnIndex);
                if (cell == null) continue;

                string raw = cell.ToString();
                if (string.IsNullOrWhiteSpace(raw))
                    continue;

                string idStr = ExtractIdNumber(raw);

                if (int.TryParse(idStr, out int idVal))
                    ids.Add(idVal);
            }

            return new HashSet<int>(ids).ToList();
        }

        private string ExtractIdNumber(string text)
        {
            List<char> digits = new List<char>();
            foreach (char c in text)
            {
                if (char.IsDigit(c))
                    digits.Add(c);
            }
            return new string(digits.ToArray());
        }
    }
}
