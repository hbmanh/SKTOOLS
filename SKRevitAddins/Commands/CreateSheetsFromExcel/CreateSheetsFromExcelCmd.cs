using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SKRevitAddins.Commands.CreateSheetsFromExcel
{
    [Transaction(TransactionMode.Manual)]
    public class CreateSheetsFromExcelCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            var dialog = new ExcelSelectionForm(doc);
            var result = dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                var saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    FileName = "SheetTemplate.xlsx"
                };

                if (saveDialog.ShowDialog() != true)
                    return Result.Cancelled;

                ExcelHelper.CreateTemplateExcel(saveDialog.FileName);
                return Result.Succeeded;
            }

            if (dialog.DialogResult != System.Windows.Forms.DialogResult.OK || string.IsNullOrEmpty(dialog.SelectedFilePath))
                return Result.Cancelled;

            string excelPath = dialog.SelectedFilePath;
            string selectedTitleBlockName = dialog.SelectedTitleBlock;

            var sheetData = ExcelHelper.ReadExcel(excelPath);
            var existingSheets = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>().ToList();
            var duplicates = new List<(string, string)>();

            foreach (var entry in sheetData)
            {
                if (existingSheets.Any(s => s.SheetNumber == entry.Value.number))
                    duplicates.Add((entry.Value.number, entry.Value.name));
            }

            HashSet<string> sheetNumbersToKeep = new();

            if (duplicates.Any())
            {
                var selForm = new SheetSelectionForm(duplicates);
                if (selForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var selectedToReplace = selForm.SelectedSheets;
                    sheetNumbersToKeep = selectedToReplace.ToHashSet();

                    using (Transaction tx = new Transaction(doc, "Xóa sheet và view trùng"))
                    {
                        tx.Start();

                        foreach (string sn in selectedToReplace)
                        {
                            var sheet = existingSheets.FirstOrDefault(s => s.SheetNumber == sn);
                            if (sheet == null) continue;

                            string sheetNumber = sheet.SheetNumber;
                            string sheetName = sheet.Name;
                            string expectedViewName = $"{sheetNumber} - {sheetName}";

                            var viewports = new FilteredElementCollector(doc)
                                .OfClass(typeof(Viewport))
                                .Cast<Viewport>()
                                .Where(vp => vp.SheetId == sheet.Id)
                                .ToList();

                            foreach (var vp in viewports)
                            {
                                try
                                {
                                    var viewElem = doc.GetElement(vp.ViewId);
                                    if (viewElem is View view &&
                                        !view.IsTemplate &&
                                        view.Name.Equals(expectedViewName, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        doc.Delete(view.Id);
                                    }

                                    doc.Delete(vp.Id);
                                }
                                catch (InvalidObjectException)
                                {
                                    continue;
                                }
                            }

                            try
                            {
                                doc.Delete(sheet.Id);
                            }
                            catch (InvalidObjectException) { }
                        }

                        tx.Commit();
                    }

                    // 🔁 Cập nhật lại danh sách sheet tồn tại sau khi xoá
                    var remainingSheetNumbers = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .Select(s => s.SheetNumber)
                        .ToHashSet();

                    sheetData = sheetData
                        .Where(kv => sheetNumbersToKeep.Contains(kv.Value.number) ||
                                     !remainingSheetNumbers.Contains(kv.Value.number))
                        .ToDictionary(kv => kv.Key, kv => kv.Value);
                }
                else
                {
                    return Result.Cancelled;
                }
            }

            var titleBlock = GetTitleBlock(doc, selectedTitleBlockName);
            if (titleBlock == null)
            {
                message = "Không tìm thấy khung tên đã chọn.";
                return Result.Failed;
            }

            var progress = new ProgressForm(sheetData.Count);
            progress.Show();

            using (Transaction tx = new Transaction(doc, "Tạo Sheets và Views"))
            {
                tx.Start();

                int count = 0;
                foreach (var entry in sheetData)
                {
                    var (sheetNumber, sheetName, viewGroup, viewFlag, levelName) = entry.Value;
                    string viewFullName = $"{sheetNumber} - {sheetName}";
                    View viewToCreate = null;

                    if (!string.IsNullOrWhiteSpace(viewFlag))
                    {
                        if (viewFlag.ToUpper() == "DV")
                        {
                            var detailType = new FilteredElementCollector(doc)
                                .OfClass(typeof(ViewFamilyType))
                                .Cast<ViewFamilyType>()
                                .FirstOrDefault(v => v.ViewFamily == ViewFamily.Drafting);

                            if (detailType != null)
                            {
                                viewToCreate = ViewDrafting.Create(doc, detailType.Id);
                                viewToCreate.Name = viewFullName;
                            }
                        }
                        else if (!string.IsNullOrWhiteSpace(levelName))
                        {
                            ViewFamily? viewFamily = viewFlag.ToUpper() switch
                            {
                                "FL" => ViewFamily.FloorPlan,
                                "SL" => ViewFamily.StructuralPlan,
                                _ => null
                            };

                            if (viewFamily.HasValue)
                            {
                                var level = new FilteredElementCollector(doc)
                                    .OfClass(typeof(Level)).Cast<Level>()
                                    .FirstOrDefault(l => l.Name.Equals(levelName, StringComparison.InvariantCultureIgnoreCase));

                                var viewType = new FilteredElementCollector(doc)
                                    .OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                                    .FirstOrDefault(v => v.ViewFamily == viewFamily.Value);

                                if (level != null && viewType != null)
                                {
                                    viewToCreate = ViewPlan.Create(doc, viewType.Id, level.Id);
                                    viewToCreate.Name = viewFullName;
                                }
                            }
                        }
                    }

                    var newSheet = ViewSheet.Create(doc, titleBlock.Id);
                    newSheet.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set(sheetNumber);
                    newSheet.get_Parameter(BuiltInParameter.SHEET_NAME).Set(sheetName);

                    var groupParam = newSheet.LookupParameter("Sub-Package");
                    if (groupParam != null && !groupParam.IsReadOnly)
                        groupParam.Set(viewGroup);

                    var viewToPlace = new FilteredElementCollector(doc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .FirstOrDefault(v => v.Name.Equals(viewFullName, StringComparison.InvariantCultureIgnoreCase));

                    if (viewToPlace != null && Viewport.CanAddViewToSheet(doc, newSheet.Id, viewToPlace.Id))
                    {
                        BoundingBoxUV bb = newSheet.Outline;
                        XYZ center = new XYZ((bb.Min.U + bb.Max.U) / 2, (bb.Min.V + bb.Max.V) / 2, 0);
                        Viewport.Create(doc, newSheet.Id, viewToPlace.Id, center);
                    }

                    progress.UpdateProgress(++count);
                }

                tx.Commit();
            }

            progress.Close();
            return Result.Succeeded;
        }

        private FamilySymbol GetTitleBlock(Document doc, string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return null;

            var parts = fullName.Split(new[] { " : " }, StringSplitOptions.None);
            if (parts.Length != 2) return null;

            string familyName = parts[0].Trim();
            string typeName = parts[1].Trim();

            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilySymbol>()
                .FirstOrDefault(tb => tb.Family.Name.Equals(familyName, StringComparison.InvariantCultureIgnoreCase)
                                   && tb.Name.Equals(typeName, StringComparison.InvariantCultureIgnoreCase));
        }
    }
}
