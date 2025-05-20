using System;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.CopySetOfFilterFromViewTemp
{
    [Transaction(TransactionMode.Manual)]
    public class CopySetOfFilterFromViewTempCmd : IExternalCommand
    {
        CopySetOfFilterFromViewTempViewModel viewModel;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            viewModel = new CopySetOfFilterFromViewTempViewModel(uiapp);

            try
            {
                App.thisApp.ShowCopySetFilterFromViewTempViewModel(uiapp, viewModel);
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