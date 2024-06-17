using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Win32;
using SKToolsAddins.ViewModel;
using Binding = Autodesk.Revit.DB.Binding;
using Document = Autodesk.Revit.DB.Document;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;

namespace SKToolsAddins.Commands.ChangeBwTypeAndIns
{
    public class ChangeBwTypeAndInsRequestHandler : IExternalEventHandler
    {
        private ChangeBwTypeAndInsViewModel ViewModel;

        public ChangeBwTypeAndInsRequestHandler(ChangeBwTypeAndInsViewModel viewModel)
        {
            ViewModel = viewModel;
        }

        private ChangeBwTypeAndInsRequest m_Request = new ChangeBwTypeAndInsRequest();

        public ChangeBwTypeAndInsRequest Request
        {
            get { return m_Request; }
        }

        public void Execute(UIApplication uiapp)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        break;
                    case (RequestId.OK):
                        ChangeBwTypeAndIns(uiapp, ViewModel);
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("エラー", ex.Message);
            }
        }

        public string GetName()
        {
            return "タイプ↔インスタンスの切り替え";
        }

        #region ChangeBwTypeAndIns

        public void ChangeBwTypeAndIns(UIApplication uiapp, ChangeBwTypeAndInsViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var selInsFamParas = viewModel.SelInsFamParasToChange;
            var selTypeFamParas = viewModel.SelTypFamParasToChange;
            var selFamFamilies = viewModel.SelFamFamilies;
            var selInsProParas = viewModel.SelInsProParasToChange;
            var selTypeProParas = viewModel.SelTypProParasToChange;
            var selProCategory = viewModel.SelProCategory;

            foreach (var selFamFamily in selFamFamilies)
            {
                Document famDoc = doc.EditFamily(selFamFamily);
                var famManager = famDoc.FamilyManager;
                using (Transaction Transaction = new Transaction(famDoc, "Change Family Param"))
                {
                    Transaction.Start();
                    foreach (var selInsFamParaInfo in selInsFamParas)
                    {
                        if (selInsFamParaInfo != null)
                        {
                            string paraInsName = selInsFamParaInfo.ParamName;
                            var paraInsType = selInsFamParaInfo.ParamType;
                            BuiltInParameterGroup paraInsGroup = selInsFamParaInfo.ParamGroup;
                            FamilyParameter famParaIns = famManager.get_Parameter(paraInsName);
                            famManager.RemoveParameter(famParaIns);
                            var newTypPara = famManager.AddParameter(paraInsName, paraInsGroup, paraInsType, false);
                        }
                    }

                    foreach (var selTypeFamPara in selTypeFamParas)
                    {
                        if (selTypeFamPara != null)
                        {
                            string paraTypeName = selTypeFamPara.ParamName;
                            var paraTypeType = selTypeFamPara.ParamType;
                            BuiltInParameterGroup paraTypeGroup = selTypeFamPara.ParamGroup;
                            FamilyParameter famParaType = famManager.get_Parameter(paraTypeName);
                            famManager.RemoveParameter(famParaType);
                            var newTypPara = famManager.AddParameter(paraTypeName, paraTypeGroup, paraTypeType, true);
                        }
                    }

                    Transaction.Commit();
                }

                MyFamilyLoadOptions familyLoadOptions = new MyFamilyLoadOptions();
                famDoc.LoadFamily(doc, familyLoadOptions);
            }

            /// Test Create File Share or Load if file share not exist

            DefinitionFile sharedParamFile;
            var sharedParamFileCheck = doc.Application.OpenSharedParameterFile();
            if (sharedParamFileCheck == null)
            {
                TaskDialog loadOrCreateDialog = new TaskDialog("Shared Param File Missing");
                loadOrCreateDialog.MainInstruction = "The shared parameter file is not found.";
                loadOrCreateDialog.MainContent = "Would you like to load an existing shared parameter file?";
                loadOrCreateDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Load an existing shared parameter file");
                loadOrCreateDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Create a new shared parameter file");
                loadOrCreateDialog.CommonButtons = TaskDialogCommonButtons.Cancel;
                loadOrCreateDialog.DefaultButton = TaskDialogResult.Cancel;
                TaskDialogResult tResult = loadOrCreateDialog.Show();

                if (tResult == TaskDialogResult.CommandLink1)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    // Prompt user to select the shared parameter file
                    openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                    openFileDialog.Title = "Select a Shared Param File";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        doc.Application.SharedParametersFilename = openFileDialog.FileName;
                    }
                }
                else if (tResult == TaskDialogResult.CommandLink2)
                {
                    Stream fileStream;
                    MemoryStream userInput = new MemoryStream();
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    // Prompt user for location to save the new shared parameter file
                    saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                    saveFileDialog.Title = "Save New Shared Param File";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Here, you might need to manually create the shared parameter file structure
                        // or use an external tool/library to create it.
                        // For now, just setting the path:
                        doc.Application.SharedParametersFilename = saveFileDialog.FileName;
                        fileStream = saveFileDialog.OpenFile();
                        userInput.Position = 0;
                        userInput.WriteTo(fileStream);
                        fileStream.Close();
                    }

                }
                sharedParamFile = doc.Application.OpenSharedParameterFile();
                using (Transaction transaction = new Transaction(doc, "Change Project Param"))
                {
                    transaction.Start();
                    foreach (var selInsProParaInfo in selInsProParas)
                    {
                        if (selInsProParaInfo != null)
                        {
                            string paraInsName = selInsProParaInfo.ParamName;
                            ParameterType paraInsType = selInsProParaInfo.ParamType;
                            BuiltInParameterGroup paraInsGroup = selInsProParaInfo.ParamGroup;
                            CategorySet categories = new CategorySet();
                            categories.Insert(selProCategory);

                            doc.Delete(selInsProParaInfo.ParamValue.Id);

                            using (SubTransaction txSubTransaction = new SubTransaction(doc))
                            {
                                txSubTransaction.Start();

                                DefinitionGroup foundGroup = FindGroupOfParameter(sharedParamFile, paraInsName);
                                if (foundGroup == null)
                                {
                                    foundGroup = GetOrCreateGroup(sharedParamFile, "カスタムパラメータ");
                                }

                                Definition definition;
                                if (foundGroup.Definitions.get_Item(paraInsName) != null)
                                {
                                    definition = foundGroup.Definitions.get_Item(paraInsName);
                                }
                                else
                                {
                                    ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(paraInsName, paraInsType);
                                    definition = foundGroup.Definitions.Create(options);
                                }

                                TypeBinding binding = doc.Application.Create.NewTypeBinding(categories);
                                doc.ParameterBindings.Insert(definition, binding, paraInsGroup);

                                txSubTransaction.Commit();
                            }

                        }
                    }

                    foreach (var selTypeProParaInfo in selTypeProParas)
                    {
                        if (selTypeProParaInfo != null)
                        {
                            string paraTypName = selTypeProParaInfo.ParamName;
                            ParameterType paraTypType = selTypeProParaInfo.ParamType;
                            BuiltInParameterGroup paraTypGroup = selTypeProParaInfo.ParamGroup;
                            CategorySet categories = new CategorySet();
                            categories.Insert(selProCategory);

                            doc.Delete(selTypeProParaInfo.ParamValue.Id);

                            using (SubTransaction txSubTransaction = new SubTransaction(doc))
                            {
                                txSubTransaction.Start();

                                DefinitionGroup foundGroup = FindGroupOfParameter(sharedParamFile, paraTypName);
                                if (foundGroup == null)
                                {
                                    foundGroup = GetOrCreateGroup(sharedParamFile, "カスタムパラメータ");
                                }


                                Definition definition;
                                if (foundGroup.Definitions.get_Item(paraTypName) != null)
                                {
                                    definition = foundGroup.Definitions.get_Item(paraTypName);
                                }
                                else
                                {
                                    ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(paraTypName, paraTypType);
                                    definition = foundGroup.Definitions.Create(options);
                                }

                                InstanceBinding binding = doc.Application.Create.NewInstanceBinding(categories);
                                doc.ParameterBindings.Insert(definition, binding, paraTypGroup);

                                txSubTransaction.Commit();
                            }
                        }
                    }
                    transaction.Commit();
                }
            }
            else
            {
                sharedParamFile = doc.Application.OpenSharedParameterFile();
                using (Transaction transaction = new Transaction(doc, "Change Project Param"))
                {
                    transaction.Start();
                    foreach (var selInsProParaInfo in selInsProParas)
                    {
                        if (selInsProParaInfo != null)
                        {
                            string paraInsName = selInsProParaInfo.ParamName;
                            ParameterType paraInsType = selInsProParaInfo.ParamType;
                            BuiltInParameterGroup paraInsGroup = selInsProParaInfo.ParamGroup;
                            CategorySet categories = new CategorySet();
                            categories.Insert(selProCategory);
                            
                            doc.Delete(selInsProParaInfo.ParamValue.Id);

                            using (SubTransaction txSubTransaction = new SubTransaction(doc))
                            {
                                txSubTransaction.Start();

                                DefinitionGroup foundGroup = FindGroupOfParameter(sharedParamFile, paraInsName);
                                if (foundGroup == null)
                                {
                                    foundGroup = GetOrCreateGroup(sharedParamFile, "カスタムパラメータ");
                                }

                                Definition definition;
                                if (foundGroup.Definitions.get_Item(paraInsName) != null)
                                {
                                    definition = foundGroup.Definitions.get_Item(paraInsName);
                                }
                                else
                                {
                                    ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(paraInsName, paraInsType);
                                    definition = foundGroup.Definitions.Create(options);
                                }

                                TypeBinding binding = doc.Application.Create.NewTypeBinding(categories);
                                doc.ParameterBindings.Insert(definition, binding, paraInsGroup);

                                txSubTransaction.Commit();
                            }

                        }
                    }

                    foreach (var selTypeProParaInfo in selTypeProParas)
                    {
                        if (selTypeProParaInfo != null)
                        {
                            string paraTypName = selTypeProParaInfo.ParamName;
                            ParameterType paraTypType = selTypeProParaInfo.ParamType;
                            BuiltInParameterGroup paraTypGroup = selTypeProParaInfo.ParamGroup;
                            CategorySet categories = new CategorySet();
                            categories.Insert(selProCategory);

                            doc.Delete(selTypeProParaInfo.ParamValue.Id);

                            using (SubTransaction txSubTransaction = new SubTransaction(doc))
                            {
                                txSubTransaction.Start();

                                DefinitionGroup foundGroup = FindGroupOfParameter(sharedParamFile, paraTypName);
                                if (foundGroup == null)
                                {
                                    foundGroup = GetOrCreateGroup(sharedParamFile, "カスタムパラメータ");
                                }


                                Definition definition;
                                if (foundGroup.Definitions.get_Item(paraTypName) != null)
                                {
                                    definition = foundGroup.Definitions.get_Item(paraTypName);
                                }
                                else
                                {
                                    ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(paraTypName, paraTypType);
                                    definition = foundGroup.Definitions.Create(options);
                                }

                                InstanceBinding binding = doc.Application.Create.NewInstanceBinding(categories);
                                doc.ParameterBindings.Insert(definition, binding, paraTypGroup);

                                txSubTransaction.Commit();
                            }
                        }
                    }
                    transaction.Commit();
                }
            }

            ///Create new project parameter not shared
            //using (Transaction transaction = new Transaction(doc, "Change Project Param"))
            //{
            //    transaction.Start();
            //    foreach (var selInsProParaInfo in selInsProParas)
            //    {
            //        if (selInsProParaInfo != null)
            //        {
            //            string paraInsName = selInsProParaInfo.ParamName;
            //            ParameterType paraInsType = selInsProParaInfo.ParamType;
            //            BuiltInParameterGroup paraInsGroup = selInsProParaInfo.ParamGroup;
            //            CategorySet categories = new CategorySet();
            //            categories.Insert(selProCategory);

            //            doc.Delete(selInsProParaInfo.ParamValue.Id);

            //            using (SubTransaction txSubTransaction = new SubTransaction(doc))
            //            {
            //                txSubTransaction.Start();

            //                var sharedParamFile = doc.Application.OpenSharedParameterFile();
            //                if (sharedParamFile == null)
            //                {
            //                    CreateProParaNotShare(doc, paraInsName, paraInsType, paraInsGroup, selProCategory, true);
            //                }
            //                else
            //                {
            //                    if (IsParameterInSharedFile(sharedParamFile, paraInsName))
            //                    {
            //                        DefinitionGroup foundGroup = FindGroupOfParameter(sharedParamFile, paraInsName);
            //                        if (foundGroup == null)
            //                        {
            //                            foundGroup = GetOrCreateGroup(sharedParamFile, "カスタムパラメータ");
            //                        }

            //                        Definition definition;
            //                        if (foundGroup.Definitions.get_Item(paraInsName) != null)
            //                        {
            //                            definition = foundGroup.Definitions.get_Item(paraInsName);
            //                        }
            //                        else
            //                        {
            //                            ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(paraInsName, paraInsType);
            //                            definition = foundGroup.Definitions.Create(options);
            //                        }

            //                        TypeBinding binding = doc.Application.Create.NewTypeBinding(categories);
            //                        doc.ParameterBindings.Insert(definition, binding, paraInsGroup);
            //                    }
            //                    else
            //                    {
            //                        CreateProParaNotShare(doc, paraInsName, paraInsType, paraInsGroup, selProCategory, true);
            //                    }
            //                }

            //                txSubTransaction.Commit();
            //            }

            //        }
            //    }

            //    foreach (var selTypeProParaInfo in selTypeProParas)
            //    {
            //        if (selTypeProParaInfo != null)
            //        {
            //            string paraTypName = selTypeProParaInfo.ParamName;
            //            ParameterType paraTypType = selTypeProParaInfo.ParamType;
            //            BuiltInParameterGroup paraTypGroup = selTypeProParaInfo.ParamGroup;
            //            CategorySet categories = new CategorySet();
            //            categories.Insert(selProCategory);

            //            doc.Delete(selTypeProParaInfo.ParamValue.Id);

            //            using (SubTransaction txSubTransaction = new SubTransaction(doc))
            //            {
            //                txSubTransaction.Start();

            //                var sharedParamFile = doc.Application.OpenSharedParameterFile();

            //                if (sharedParamFile == null)
            //                {
            //                    CreateProParaNotShare(doc, paraTypName, paraTypType, paraTypGroup, selProCategory, false);
            //                }
            //                else
            //                {
            //                    if (IsParameterInSharedFile(sharedParamFile, paraTypName))
            //                    {
            //                        DefinitionGroup foundGroup = FindGroupOfParameter(sharedParamFile, paraTypName);
            //                        if (foundGroup == null)
            //                        {
            //                            foundGroup = GetOrCreateGroup(sharedParamFile, "CustomParameters");
            //                        }


            //                        Definition definition;
            //                        if (foundGroup.Definitions.get_Item(paraTypName) != null)
            //                        {
            //                            definition = foundGroup.Definitions.get_Item(paraTypName);
            //                        }
            //                        else
            //                        {
            //                            ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(paraTypName, paraTypType);
            //                            definition = foundGroup.Definitions.Create(options);
            //                        }

            //                        InstanceBinding binding = doc.Application.Create.NewInstanceBinding(categories);
            //                        doc.ParameterBindings.Insert(definition, binding, paraTypGroup);
            //                    }
            //                    else
            //                    {
            //                        CreateProParaNotShare(doc, paraTypName, paraTypType, paraTypGroup, selProCategory, false);
            //                    }
            //                }


            //                txSubTransaction.Commit();
            //            }
            //        }
            //    }
            //    transaction.Commit();
            //}


            DefinitionGroup FindGroupOfParameter(DefinitionFile sharedParaFile, string paramName)
            {
                foreach (DefinitionGroup group in sharedParaFile.Groups)
                {
                    if (group.Definitions.get_Item(paramName) != null)
                    {
                        return group;
                    }
                }
                return null;
            }
        }

        #endregion

        public class MyFamilyLoadOptions : IFamilyLoadOptions
        {
            public bool OnFamilyFound(
                bool familyInUse,
                out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true;
            }

            public bool OnSharedFamilyFound(
                Family sharedFamily,
                bool familyInUse,
                out FamilySource source,
                out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = true;

                return true;
            }
        }
        private DefinitionGroup GetOrCreateGroup(DefinitionFile sharedParamFile, string groupName)
        {
            foreach (DefinitionGroup group in sharedParamFile.Groups)
            {
                if (group.Name == groupName)
                {
                    return group;
                }
            }
            return sharedParamFile.Groups.Create(groupName);
        }
        /// <summary>
        /// Create Project Param Not Shared
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="paramName"></param>
        /// <param name="paramType"></param>
        /// <param name="paramGroup"></param>
        /// <param name="category"></param>
        /// <param name="isInstance"></param>
        //public void CreateProParaNotShare(Document doc, string paramName, ParameterType paramType, BuiltInParameterGroup paramGroup, Category category, bool isInstance)
        //{
        //    using (Transaction trans = new Transaction(doc, "Create Project Param"))
        //    {
        //        trans.Start();

        //        CategorySet categorySet = new CategorySet();
        //        categorySet.Insert(category);

        //        ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(paramName, paramType);
        //        Definition definition = new Definitions().Create(options); // Create a new definition

        //        Binding binding;
        //        if (isInstance)
        //        {
        //            binding = doc.Application.Create.NewInstanceBinding(categorySet);
        //        }
        //        else
        //        {
        //            binding = doc.Application.Create.NewTypeBinding(categorySet);
        //        }

        //        bool success = doc.ParameterBindings.Insert(definition, binding, paramGroup);

        //        if (!success)
        //        {
        //            TaskDialog.Show("Error", "Failed to create the project parameter.");
        //        }

        //        trans.Commit();
        //    }
        //}

        /// <summary>
        /// Is Param Exits In Shared File
        /// </summary>
        /// <param name="sharedParamFile"></param>
        /// <param name="paramName"></param>
        /// <returns></returns>
        //public bool IsParameterInSharedFile(DefinitionFile sharedParamFile, string paramName)
        //{
        //    // Iterate through all groups in the Shared Param File
        //    foreach (DefinitionGroup group in sharedParamFile.Groups)
        //    {
        //        // Check if the definition exists in the group
        //        if (group.Definitions.get_Item(paramName) != null)
        //        {
        //            return true; // If found, return true
        //        }
        //    }
        //    return false; // If not found, return false
        //}

    }
}
