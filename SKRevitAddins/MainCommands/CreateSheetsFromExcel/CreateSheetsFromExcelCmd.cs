using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.CreateSheetsFromExcel
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
            bool createSheets = dialog.CreateSheets;
            bool createSheetViews = dialog.CreateSheetViews;
            bool createWorkingView = dialog.CreateWorkingView;
            string workingViewSuffix = dialog.WorkingViewSuffix;

            var sheetData = ExcelHelper.ReadExcel(excelPath);

            // Get existing sheets
            var existingSheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .GroupBy(s => s.SheetNumber)
                .ToDictionary(g => g.Key, g => g.First());

            // Fix: Create view dictionary with duplicate key checking
            var viewDict = new Dictionary<string, View>(StringComparer.InvariantCultureIgnoreCase);
            foreach (var v in new FilteredElementCollector(doc)
                         .OfClass(typeof(View))
                         .Cast<View>()
                         .Where(v => !v.IsTemplate && !string.IsNullOrWhiteSpace(v.Name)))
            {
                string key = v.Name.Trim();
                if (!string.IsNullOrWhiteSpace(key))
                {
                    if (!viewDict.ContainsKey(key))
                        viewDict.Add(key, v);
                }
            }
            var existingViews = viewDict;

            var titleBlock = GetTitleBlock(doc, selectedTitleBlockName);
            if (titleBlock == null)
            {
                message = "Không tìm thấy khung tên đã chọn.";
                return Result.Failed;
            }

            var progress = new ProgressForm(sheetData.Count);
            progress.Show();

            // Fix: Create viewTypes dictionary with duplicate key checking
            var viewTypesDict = new Dictionary<ViewFamily, ViewFamilyType>();
            foreach (var vft in new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>())
            {
                if (!viewTypesDict.ContainsKey(vft.ViewFamily))
                    viewTypesDict.Add(vft.ViewFamily, vft);
            }
            var viewTypes = viewTypesDict;

            // Fix: Create levels dictionary with duplicate key checking
            var levelsDict = new Dictionary<string, Level>(StringComparer.InvariantCultureIgnoreCase);
            foreach (var level in new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>())
            {
                string key = level.Name;
                if (!string.IsNullOrWhiteSpace(key) && !levelsDict.ContainsKey(key))
                    levelsDict.Add(key, level);
            }
            var levels = levelsDict;

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
                    progress.UpdateProgress(++count, viewFullName);

                    bool sheetExists = existingSheets.ContainsKey(sheetNumber);
                    bool viewExists = existingViews.ContainsKey(viewFullName);

                    View viewToPlace = null;

                    if (createSheetViews && !viewExists && !string.IsNullOrWhiteSpace(viewFlag))
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

                    if (createWorkingView && !string.IsNullOrWhiteSpace(workingViewSuffix) && viewToPlace != null)
                    {
                        string workingViewName = $"{sheetNumber} - {sheetName}_{workingViewSuffix}";
                        if (!existingViews.ContainsKey(workingViewName))
                        {
                            View workingView = null;

                            if (viewToPlace is ViewDrafting && viewTypes.TryGetValue(ViewFamily.Drafting, out var draftingType))
                            {
                                workingView = ViewDrafting.Create(doc, draftingType.Id);
                            }
                            else if (viewToPlace is ViewPlan && !string.IsNullOrWhiteSpace(viewFlag))
                            {
                                var vf = viewFlag.ToUpper();
                                var family = vf == "FL" ? ViewFamily.FloorPlan : ViewFamily.StructuralPlan;

                                if (viewTypes.TryGetValue(family, out var viewType) &&
                                    levels.TryGetValue(levelName, out var level))
                                {
                                    workingView = ViewPlan.Create(doc, viewType.Id, level.Id);
                                }
                            }

                            if (workingView != null)
                            {
                                workingView.Name = workingViewName;
                                existingViews[workingViewName] = workingView;
                            }
                        }
                    }
                    // Set "View Type" parameter to "02 Sheet Views"
                    Parameter viewTypeParam = viewToPlace.LookupParameter("View Type");
                    if (viewTypeParam != null && !viewTypeParam.IsReadOnly)
                    {
                        viewTypeParam.Set("02 Sheet Views");
                    }
                    // Set "View Sub-Type" parameter to match the viewGroup value
                    Parameter viewSubTypeParam = viewToPlace.LookupParameter("View Sub-Type");
                    if (viewSubTypeParam != null && !viewSubTypeParam.IsReadOnly)
                    {
                        viewSubTypeParam.Set(viewGroup);
                    }
                    if (createSheets && !sheetExists)
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
                .FirstOrDefault(tb =>
                    tb.Family.Name.Equals(familyName, StringComparison.InvariantCultureIgnoreCase)
                    && tb.Name.Equals(typeName, StringComparison.InvariantCultureIgnoreCase));
        }
    }
}