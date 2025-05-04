using System;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKRevitAddins.ViewModel;

namespace SKRevitAddins.PermissibleRangeFrame
{
    [Transaction(TransactionMode.Manual)]
    public class PermissibleRangeFrameCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var viewModel = new PermissibleRangeFrameViewModel(uiapp);

            try
            {
                // Giả sử App.thisApp.ShowPermissibleRangeFrameViewModel là phương thức hiển thị UI (cửa sổ WPF)
                App.thisApp.ShowPermissibleRangeFrameViewModel(uiapp, viewModel);
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
