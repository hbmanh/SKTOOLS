using System;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKRevitAddins.ViewModel;

namespace SKRevitAddins.SelectElements
{
    [Transaction(TransactionMode.Manual)]
    public class SelectElementsCmd : IExternalCommand
    {
        SelectElementsViewModel viewModel;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            viewModel = new SelectElementsViewModel(uiapp);

            try
            {
                App.thisApp.ShowSelectElementsViewModel(uiapp, viewModel);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}