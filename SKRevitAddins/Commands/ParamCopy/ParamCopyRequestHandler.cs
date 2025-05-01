using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ParamCopy
{
    public class ParamCopyRequestHandler : IExternalEventHandler
    {
        public ParamCopyRequestHandler(ParamCopyViewModel viewModel)
        {
            ViewModel = viewModel;
        }
        private ParamCopyViewModel ViewModel;
        private ParamCopyRequest m_Request = new ParamCopyRequest();
        public ParamCopyRequest Request
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
                    case RequestId.InstanceCopy:
                        InstanceCopy(uiapp, ViewModel);
                        break;
                    case RequestId.FamilyCopy:
                        FamilyCopy(uiapp, ViewModel);
                        break;
                    case RequestId.CategoryCopy:
                        CategoryCopy(uiapp, ViewModel);
                        break;
                    case RequestId.AllEleCopy:
                        AllEleCopy(uiapp, ViewModel);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                // TaskDialog.Show("エラー", ex.Message);
                MessageBox.Show(ex.Message, "エラー");
            }
        }

        public string GetName()
        {
            return "パラメーターコピー";
        }
        
        private void InstanceCopy(UIApplication uiapp, ParamCopyViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            Element element = null;
            if (viewModel.IsTargetInstTabEnabled)
                element = viewModel.TargetElement;
            if (viewModel.IsTargetTypeTabEnabled)
                element = viewModel.TargetElementType;

            try
            {
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("パラメーターコピー");

                    element.get_Parameter(viewModel.SelectedTargetParam.Param.Definition)?.Set(viewModel.SourceValue);

                    tx.Commit();
                }
            }
            catch
            {
            }
        }
        private void FamilyCopy(UIApplication uiapp, ParamCopyViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (viewModel.SameFamilyTargetElements.Count <= 0) return;

            try
            {
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("パラメーターコピー");
                    foreach (var ele in viewModel.SameFamilyTargetElements)
                    {
                        Element element = null;
                        if (viewModel.IsTargetInstTabEnabled)
                            element = ele;
                        if (viewModel.IsTargetTypeTabEnabled)
                            element = ele.GetElementType(doc);

                        element.get_Parameter(viewModel.SelectedTargetParam.Param.Definition)?.Set(viewModel.SourceValue);
                    }
                    tx.Commit();
                }
                
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CategoryCopy(UIApplication uiapp, ParamCopyViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (viewModel.SameCategoryTargetElements.Count <= 0) return;

            try
            {
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("パラメーターコピー");
                    foreach (var ele in viewModel.SameCategoryTargetElements)
                    {
                        Element element = null;
                        if (viewModel.IsTargetInstTabEnabled)
                            element = ele;
                        if (viewModel.IsTargetTypeTabEnabled)
                            element = ele.GetElementType(doc);

                        element.get_Parameter(viewModel.SelectedTargetParam.Param.Definition)?.Set(viewModel.SourceValue);
                    }
                    tx.Commit();
                }
                
            }
            catch
            {
            }
        }
        private void AllEleCopy(UIApplication uiapp, ParamCopyViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var allEle = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .ToList();

            try
            {
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("パラメーターコピー");
                    foreach (var ele in allEle)
                    {
                        Element element = null;
                        if (viewModel.IsTargetInstTabEnabled)
                            element = ele;
                        if (viewModel.IsTargetTypeTabEnabled)
                            element = ele.GetElementType(doc);
                        if (element == null) continue;
                        element.get_Parameter(viewModel.SelectedTargetParam.Param.Definition)?.Set(viewModel.SourceValue);
                    }
                    tx.Commit();
                }
            }
            catch
            {
            }
        }
    }
}
