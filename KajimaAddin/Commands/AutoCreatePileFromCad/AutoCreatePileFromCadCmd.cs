using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using Autodesk.Revit.ApplicationServices;
using SKToolsAddins.ViewModel;

namespace SKToolsAddins.Commands.AutoCreatePileFromCad
{
    [Transaction(TransactionMode.Manual)]
    public class AutoCreatePileFromCadCmd : IExternalCommand
    {
        AutoCreatePileFromCadViewModel viewModel;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            viewModel = new AutoCreatePileFromCadViewModel(uiapp);

            try
            {
                App.thisApp.ShowAutoCreatePileFromCadViewModel(uiapp, viewModel);
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
