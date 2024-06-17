using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using Autodesk.Revit.ApplicationServices;
using SKToolsAddins.ViewModel;

namespace SKToolsAddins.Commands.FindDWGNotUsedAndDel
{
    [Transaction(TransactionMode.Manual)]
    public class FindDWGNotUsedAndDelCmd : IExternalCommand
    {
        FindDWGNotUsedAndDelViewModel viewModel;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            viewModel = new FindDWGNotUsedAndDelViewModel(uiapp);

            try
            {
                App.thisApp.ShowFindDWGNotUseAndDelViewModel(uiapp, viewModel);
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