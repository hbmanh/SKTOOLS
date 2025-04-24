using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
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
            var existingSheets = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>().ToDictionary(s => s.SheetNumber);
            var existingViews = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>()
                .Where(v => !v.IsTemplate && !string.IsNullOrWhiteSpace(v.Name))
                .ToDictionary(v => v.Name, v => v, StringComparer.InvariantCultureIgnoreCase);

            var titleBlock = GetTitleBlock(doc, selectedTitleBlockName);
            if (titleBlock == null)
            {
                message = "Không tìm thấy khung tên đã chọn.";
                return Result.Failed;
            }

            var progress = new ProgressForm(sheetData.Count);
            progress.Show();

            var viewTypes = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().ToDictionary(v => v.ViewFamily);
            var levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToDictionary(l => l.Name, StringComparer.InvariantCultureIgnoreCase);

            using (Transaction tx = new Transaction(doc, "Tạo Sheet và View"))
            {
                tx.Start();

                int count = 0;
                foreach (var entry in sheetData)
                {
                    if (progress.IsCanceled)
                    {
                        tx.RollBack();
                        progress.Close();
                        TaskDialog.Show("Đã huỷ", "Toàn bộ thao tác đã được huỷ.");
                        return Result.Cancelled;
                    }

                    var (sheetNumber, sheetName, viewGroup, viewFlag, levelName) = entry.Value;
                    string viewFullName = $"{sheetNumber} - {sheetName}";

                    progress.UpdateProgress(++count, $"{sheetNumber} - {sheetName}");

                    bool sheetExists = existingSheets.ContainsKey(sheetNumber);
                    bool viewExists = existingViews.ContainsKey(viewFullName);

                    View viewToPlace = null;

                    // ❗ Nếu view chưa tồn tại, tạo mới
                    if (!viewExists && !string.IsNullOrWhiteSpace(viewFlag))
                    {
                        string vf = viewFlag.ToUpper();
                        if (vf == "DV" && viewTypes.TryGetValue(ViewFamily.Drafting, out var draftingType))
                        {
                            viewToPlace = ViewDrafting.Create(doc, draftingType.Id);
                        }
                        else if ((vf == "FL" || vf == "SL") && !string.IsNullOrWhiteSpace(levelName))
                        {
                            var family = vf == "FL" ? ViewFamily.FloorPlan : ViewFamily.StructuralPlan;
                            if (levels.TryGetValue(levelName, out var level) &&
                                viewTypes.TryGetValue(family, out var viewType))
                            {
                                viewToPlace = ViewPlan.Create(doc, viewType.Id, level.Id);
                            }
                        }

                        if (viewToPlace != null)
                        {
                            viewToPlace.Name = viewFullName;
                            existingViews[viewFullName] = viewToPlace;
                        }
                    }
                    else if (viewExists)
                    {
                        viewToPlace = existingViews[viewFullName];
                    }

                    // ❗ Nếu sheet chưa tồn tại → tạo sheet và đặt view (nếu có)
                    if (!sheetExists)
                    {
                        var newSheet = ViewSheet.Create(doc, titleBlock.Id);
                        newSheet.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set(sheetNumber);
                        newSheet.get_Parameter(BuiltInParameter.SHEET_NAME).Set(sheetName);

                        var groupParam = newSheet.LookupParameter("Sub-Package");
                        if (groupParam != null && !groupParam.IsReadOnly)
                            groupParam.Set(viewGroup);

                        if (viewToPlace != null && Viewport.CanAddViewToSheet(doc, newSheet.Id, viewToPlace.Id))
                        {
                            BoundingBoxUV bb = newSheet.Outline;
                            XYZ center = new XYZ((bb.Min.U + bb.Max.U) / 2, (bb.Min.V + bb.Max.V) / 2, 0);
                            Viewport.Create(doc, newSheet.Id, viewToPlace.Id, center);
                        }

                        existingSheets[sheetNumber] = newSheet;
                    }
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
