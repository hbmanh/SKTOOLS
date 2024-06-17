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
            string excelFilePath1 = GetExcelFilePath("Select the first Excel file (竣工BIM01-エレメント表)");
            if (string.IsNullOrEmpty(excelFilePath1))
            {
                message = "No file selected.";
                return Result.Failed;
            }

            // Show file dialog to select the second Excel file
            string excelFilePath2 = GetExcelFilePath("Select the second Excel file (竣工BIM02-BCI)");
            if (string.IsNullOrEmpty(excelFilePath2))
            {
                message = "No file selected.";
                return Result.Failed;
            }

            // Read data from Excel files and create ParamObj list
            List<ParamObj> parameterObjs = ReadParametersFromExcel(excelFilePath1, excelFilePath2, doc);

            // Process the parameterObjs and assign shared parameters
            return ProcessParameters(doc, app, parameterObjs, ref message);
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
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
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
                                string categoryName = sheet1.Cells[4, col].Text; // Read from row 5
                                if (string.IsNullOrWhiteSpace(categoryName))
                                {
                                    continue; // Skip if categoryName is empty
                                }
                                var categories = GetCategoriesFromSheet2(sheet2, categoryName, doc);
                                paramObj.Categories.AddRange(categories);
                            }
                        }

                        parameterObjs.Add(paramObj);
                    }
                }
            }

            return parameterObjs;
        }

        private List<Category> GetCategoriesFromSheet2(ExcelWorksheet sheet, string categoryName, Document doc)
        {
            List<Category> categories = new List<Category>();

            for (int row = 5; row <= sheet.Dimension.End.Row; row++)
            {
                if (IsStrikethrough(sheet.Cells[row, 3]) || string.IsNullOrWhiteSpace(sheet.Cells[row, 3].Text))
                {
                    continue; // Skip strikethrough and empty rows
                }

                if (sheet.Cells[row, 2].Text == categoryName)
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
                message = "Shared parameter file is not set.";
                return Result.Failed;
            }

            using (Transaction transaction = new Transaction(doc, "Assign Shared Parameters"))
            {
                transaction.Start();

                foreach (var paramObj in parameterObjs)
                {
                    string paramName = SanitizeParameterName(paramObj.ParamName);
                    DefinitionGroup sharedParamGroup = sharedParamFile.Groups.FirstOrDefault(g => g.Name == "CustomParameters")
                                                       ?? sharedParamFile.Groups.Create("CustomParameters");
                    Definition definition = sharedParamGroup.Definitions.FirstOrDefault(d => d.Name == paramName);

                    if (definition == null)
                    {
                        ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.String.Text);
                        definition = sharedParamGroup.Definitions.Create(options);
                    }
                    else
                    {
                        // If parameter already exists in shared parameter file, skip the creation
                        definition = sharedParamGroup.Definitions.FirstOrDefault(d => d.Name == paramName);
                    }

                    // Check if the parameter already exists in the project
                    BindingMap bindingMap = doc.ParameterBindings;
                    if (bindingMap.Contains(definition))
                    {
                        continue; // Skip if parameter already exists in the project
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
                TaskDialog defineSharedParamFileDialog = new TaskDialog("Shared Parameter File")
                {
                    MainInstruction = "No shared parameter file is set. Please select or create a shared parameter file.",
                    MainContent = "Would you like to select an existing shared parameter file or create a new one?",
                    CommonButtons = TaskDialogCommonButtons.Cancel,
                    DefaultButton = TaskDialogResult.Cancel
                };
                defineSharedParamFileDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Select existing shared parameter file");
                defineSharedParamFileDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Create new shared parameter file");

                TaskDialogResult result = defineSharedParamFileDialog.Show();

                switch (result)
                {
                    case TaskDialogResult.CommandLink1:
                        OpenFileDialog openFileDialog = new OpenFileDialog
                        {
                            Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                            Title = "Select Shared Parameter File"
                        };

                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            app.SharedParametersFilename = openFileDialog.FileName;
                        }
                        break;

                    case TaskDialogResult.CommandLink2:
                        SaveFileDialog saveFileDialog = new SaveFileDialog
                        {
                            Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                            Title = "Create New Shared Parameter File"
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
                        message = "Operation cancelled by user.";
                        return false;
                }
            }

            sharedParamFile = app.OpenSharedParameterFile();
            if (sharedParamFile == null)
            {
                message = "Failed to set shared parameter file.";
                return false;
            }

            return true;
        }
    }
}
