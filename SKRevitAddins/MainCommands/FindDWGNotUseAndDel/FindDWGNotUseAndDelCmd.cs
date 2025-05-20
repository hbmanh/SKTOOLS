using System;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.FindDWGNotUseAndDel
{
    [Transaction(TransactionMode.Manual)]
    public class FindDWGNotUseAndDelCmd : IExternalCommand
    {
        FindDWGNotUseAndDelViewModel viewModel;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Application app = uiApp.Application;
            Document doc = uiDoc.Document;

            viewModel = new FindDWGNotUseAndDelViewModel(uiApp);

            try
            {
                // Gọi hàm ShowFindDWGNotUseAndDelViewModel trong App.cs 
                // => hiển thị cửa sổ
                App.thisApp.ShowFindDWGNotUseAndDelViewModel(uiApp, viewModel);

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
