using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Binding = Autodesk.Revit.DB.Binding;

namespace SKToolsAddins.Commands.ParameterAssignment
{
    [Transaction(TransactionMode.Manual)]
    public class CreateSharedParamFromExcelCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            Application app = uiApp.Application;

            // Set LicenseContext for ExcelPackage
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Show file dialog to select the first Excel file
            string excelFilePath1 = GetExcelFilePath("最初のExcelファイルを選択 (竣工BIM01-エレメント表)");
            if (string.IsNullOrEmpty(excelFilePath1))
            {
                message = "ファイルが選択されていません。";
                return Result.Failed;
            }

            // Show file dialog to select the second Excel file
            string excelFilePath2 = GetExcelFilePath("二番目のExcelファイルを選択 (竣工BIM02-BCI)");
            if (string.IsNullOrEmpty(excelFilePath2))
            {
                message = "ファイルが選択されていません。";
                return Result.Failed;
            }

            // Read data from Excel file and create ParamObj list
            List<ParamObj> parameterObjs = ReadParametersFromExcel(excelFilePath1, excelFilePath2, doc);

            // Process the parameterObjs and assign shared parameters
            Result result = ProcessParameters(doc, app, parameterObjs, ref message);
            if (result != Result.Succeeded)
            {
                return result;
            }

            // Create schedules for each category with corresponding parameters as fields
            return CreateSchedules(doc, parameterObjs, ref message);
        }

        public class ParamObj
        {
            public string ParamName { get; set; }
            public List<Category> Categories { get; set; }
            public bool IsInstance { get; set; }

            public ParamObj(string paramName, bool isInstance)
            {
                ParamName = paramName;
                IsInstance = isInstance;
                Categories = new List<Category>();
            }
        }

        private string GetExcelFilePath(string title)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excelファイル|*.xls;*.xlsx;*.xlsm",
                Title = title
            })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }

            return null;
        }

        private List<ParamObj> ReadParametersFromExcel(string filePath1, string filePath2, Document doc)
        {
            List<ParamObj> parameterObjs = new List<ParamObj>();

            using (FileStream stream1 = new FileStream(filePath1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage package1 = new ExcelPackage(stream1))
            {
                ExcelWorksheet sheet2 = package1.Workbook.Worksheets[1]; // Sheet 2

                using (FileStream stream2 = new FileStream(filePath2, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ExcelPackage package2 = new ExcelPackage(stream2))
                {
                    ExcelWorksheet sheet1 = package2.Workbook.Worksheets[0]; // Sheet 1

                    for (int row = 6; row <= sheet1.Dimension.End.Row; row++)
                    {
                        string parameterName = sheet1.Cells[row, 3].Text;
                        bool isInstance = sheet1.Cells[row, 1].Text == "I";
                        ParamObj paramObj = new ParamObj(parameterName, isInstance);

                        for (int col = 4; col <= 33; col++) // Hardcoded to 33
                        {
                            string cellValue = sheet1.Cells[row, col].Text;
                            if (cellValue == "O" || cellValue == "Y")
                            {
                                string categoryName = sheet1.Cells[5, col].Text; // Read from row 5
                                if (string.IsNullOrWhiteSpace(categoryName))
                                {
                                    continue; // Skip if categoryName is empty
                                }
                                var categories = GetCategoriesFromBIM01(sheet2, categoryName, doc);
                                paramObj.Categories.AddRange(categories);
                            }
                        }

                        parameterObjs.Add(paramObj);
                    }
                }
            }

            return parameterObjs;
        }

        private List<Category> GetCategoriesFromBIM01(ExcelWorksheet sheet, string categoryName, Document doc)
        {
            List<Category> categories = new List<Category>();

            for (int row = 5; row <= sheet.Dimension.End.Row; row++)
            {
                if (IsStrikethrough(sheet.Cells[row, 3]) || string.IsNullOrWhiteSpace(sheet.Cells[row, 3].Text))
                {
                    continue; // Skip strikethrough and empty rows
                }

                if (sheet.Cells[row, 3].Text == categoryName)
                {
                    string categoryInternalNames = sheet.Cells[row, 8].Text; // Column H (zero-based index 7)

                    string[] categoryNames;
                    if (categoryInternalNames.Contains('、'))
                    {
                        categoryNames = categoryInternalNames.Split('、'); // Split by delimiter '、'
                    }
                    else
                    {
                        categoryNames = new string[] { categoryInternalNames }; // Single category
                    }

                    foreach (var name in categoryNames)
                    {
                        var category = GetCategoryByName(doc, name.Trim());
                        if (category != null)
                        {
                            categories.Add(category);
                        }
                    }

                    break;
                }
            }

            return categories;
        }

        private bool IsStrikethrough(ExcelRange cell)
        {
            // Check if the cell text is strikethrough
            return cell.Style.Font.Strike;
        }

        private Category GetCategoryByName(Document doc, string categoryName)
        {
            foreach (Category category in doc.Settings.Categories)
            {
                if (category.Name.Equals(categoryName, StringComparison.OrdinalIgnoreCase))
                {
                    return category;
                }
            }

            return null;
        }

        private Result ProcessParameters(Document doc, Application app, List<ParamObj> parameterObjs, ref string message)
        {
            if (!EnsureSharedParameterFileIsSet(app, ref message))
            {
                return Result.Failed;
            }

            DefinitionFile sharedParamFile = app.OpenSharedParameterFile();
            if (sharedParamFile == null)
            {
                message = "共有パラメータファイルが設定されていません。";
                return Result.Failed;
            }

            using (Transaction transaction = new Transaction(doc, "共有パラメータを割り当てる"))
            {
                transaction.Start();

                foreach (var paramObj in parameterObjs)
                {
                    string paramName = SanitizeParameterName(paramObj.ParamName);
                    DefinitionGroup sharedParamGroup = sharedParamFile.Groups.FirstOrDefault(g => g.Name == "カスタムパラメータ")
                                                               ?? sharedParamFile.Groups.Create("カスタムパラメータ");
                    Definition definition = sharedParamGroup.Definitions.FirstOrDefault(d => d.Name == paramName);

                    if (definition == null)
                    {
                        ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.String.Text);
                        definition = sharedParamGroup.Definitions.Create(options);
                    }

                    // Check if the parameter already exists in the project
                    BindingMap bindingMap = doc.ParameterBindings;
                    Definition existingDefinition = null;
                    DefinitionBindingMapIterator iterator = bindingMap.ForwardIterator();
                    while (iterator.MoveNext())
                    {
                        if (iterator.Key.Name == paramName)
                        {
                            existingDefinition = iterator.Key as Definition;
                            break;
                        }
                    }

                    // If the parameter exists, remove it before creating a new one
                    if (existingDefinition != null)
                    {
                        bindingMap.Remove(existingDefinition);
                    }

                    CategorySet categorySet = app.Create.NewCategorySet();
                    foreach (Category category in paramObj.Categories)
                    {
                        categorySet.Insert(category);
                    }

                    if (!categorySet.IsEmpty)
                    {
                        Binding binding;
                        if (paramObj.IsInstance)
                        {
                            binding = app.Create.NewInstanceBinding(categorySet);
                        }
                        else
                        {
                            binding = app.Create.NewTypeBinding(categorySet);
                        }

                        doc.ParameterBindings.Insert(definition, binding, BuiltInParameterGroup.PG_TEXT);
                    }
                }

                transaction.Commit();
            }

            return Result.Succeeded;
        }


        private string SanitizeParameterName(string paramName)
        {
            // Remove non-printable characters
            return new string(paramName.Where(c => !char.IsControl(c)).ToArray());
        }

        private bool EnsureSharedParameterFileIsSet(Application app, ref string message)
        {
            DefinitionFile sharedParamFile = app.OpenSharedParameterFile();
            if (sharedParamFile == null)
            {
                TaskDialog defineSharedParamFileDialog = new TaskDialog("共有パラメータファイル")
                {
                    MainInstruction = "共有パラメータファイルが設定されていません。既存の共有パラメータファイルを選択するか、新しい共有パラメータファイルを作成してください。",
                    MainContent = "既存の共有パラメータファイルを選択しますか、新しい共有パラメータファイルを作成しますか？",
                    CommonButtons = TaskDialogCommonButtons.Cancel,
                    DefaultButton = TaskDialogResult.Cancel
                };
                defineSharedParamFileDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "既存の共有パラメータファイルを選択する");
                defineSharedParamFileDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "新しい共有パラメータファイルを作成する");

                TaskDialogResult result = defineSharedParamFileDialog.Show();

                switch (result)
                {
                    case TaskDialogResult.CommandLink1:
                        OpenFileDialog openFileDialog = new OpenFileDialog
                        {
                            Filter = "テキストファイル (*.txt)|*.txt|すべてのファイル (*.*)|*.*",
                            Title = "共有パラメータファイルを選択する"
                        };

                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            app.SharedParametersFilename = openFileDialog.FileName;
                        }
                        break;

                    case TaskDialogResult.CommandLink2:
                        SaveFileDialog saveFileDialog = new SaveFileDialog
                        {
                            Filter = "テキストファイル (*.txt)|*.txt|すべてのファイル (*.*)|*.*",
                            Title = "新しい共有パラメータファイルを作成する"
                        };

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string fileName = saveFileDialog.FileName;
                            if (File.Exists(fileName))
                            {
                                File.Delete(fileName);
                            }

                            File.Create(fileName).Close();
                            app.SharedParametersFilename = fileName;
                        }
                        break;

                    case TaskDialogResult.Cancel:
                        message = "ユーザーによって操作がキャンセルされました。";
                        return false;
                }
            }

            sharedParamFile = app.OpenSharedParameterFile();
            if (sharedParamFile == null)
            {
                message = "共有パラメータファイルの設定に失敗しました。";
                return false;
            }

            return true;
        }

        private Result CreateSchedules(Document doc, List<ParamObj> parameterObjs, ref string message)
        {
            using (Transaction transaction = new Transaction(doc, "スケジュールを作成"))
            {
                transaction.Start();

                var categoryParamGroups = new Dictionary<Category, List<ParamObj>>();
                foreach (var paramObj in parameterObjs)
                {
                    foreach (var category in paramObj.Categories)
                    {
                        if (!categoryParamGroups.ContainsKey(category))
                        {
                            categoryParamGroups[category] = new List<ParamObj>();
                        }
                        if (!categoryParamGroups[category].Contains(paramObj))
                        {
                            categoryParamGroups[category].Add(paramObj);
                        }
                    }
                }
                List<ParamObj> uniqueParameters = new List<ParamObj>();
                
                foreach (var categoryGroup in categoryParamGroups)
                {
                    var paramObjs = categoryGroup.Value;
                    foreach (var paramObj in paramObjs)
                    {
                        if (uniqueParameters .All(p => p.ParamName != paramObj.ParamName))
                        {
                            uniqueParameters .Add(paramObj);
                        }
                    }
                    
                }

                var parameterCategoryGroups = new Dictionary<ParamObj, List<Category>>();
                foreach (var paramObj in parameterObjs)
                {
                    foreach (var category in paramObj.Categories)
                    {
                        if (!parameterCategoryGroups.ContainsKey(paramObj))
                        {
                            parameterCategoryGroups[paramObj] = new List<Category>();
                        }
                        if (!parameterCategoryGroups[paramObj].Contains(category))
                        {
                            parameterCategoryGroups[paramObj].Add(category);
                        }
                    }
                }
                List<Category> uniqueCategories = new List<Category>();

                foreach (var parameterGroup in parameterCategoryGroups)
                {
                    var categoriesContainParamObj = parameterGroup.Value;
                    foreach (var category in categoriesContainParamObj)
                    {
                        if (uniqueCategories .All(c => c.Id != category.Id) && HasInstances(doc, category))
                        { 
                            uniqueCategories .Add(category);
                        }
                    }
                }
                foreach (var category in uniqueCategories)
                {
                    // Create a schedule for the category
                    ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, category.Id);
                    var schedulableFields = schedule.Definition.GetSchedulableFields();

                    foreach (var paramObj in uniqueParameters)
                    {
                        // Check if the parameter is bound to the category
                        bool isParameterBoundToCategory = paramObj.Categories.Any(c => c.Id == category.Id);

                        if (isParameterBoundToCategory )
                        {
                            var field = schedulableFields.FirstOrDefault(f => f.GetName(doc).Equals(paramObj.ParamName, StringComparison.OrdinalIgnoreCase));
                            // Add fields to the schedule
                            if (field != null)
                            {
                                schedule.Definition.AddField(field);
                            }
                            //if (field != null && field.ParameterId != ElementId.InvalidElementId)
                            //{
                            //    Element parameterElement = doc.GetElement(field.ParameterId);
                            //    if (parameterElement != null && parameterElement is SharedParameterElement)
                            //    {
                            //        // Add fields to the schedule
                            //        schedule.Definition.AddField(field);
                            //    }
                            //}
                        }
                    }
                }

                transaction.Commit();
            }

            return Result.Succeeded;
        }
        private bool HasInstances(Document doc, Category category)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfCategoryId(category.Id).WhereElementIsNotElementType();
            return collector.Any();
        }
    }
}
