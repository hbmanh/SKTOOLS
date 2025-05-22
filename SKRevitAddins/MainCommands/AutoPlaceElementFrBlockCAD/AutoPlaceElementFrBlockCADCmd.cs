using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
{
    [Transaction(TransactionMode.Manual)]
    public class AutoPlaceElementFrBlockCADCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 1. Tạo ViewModel
            var viewModel = new AutoPlaceElementFrBlockCADViewModel(uiapp);

            // 2. Khởi tạo handler và Window (window cần handler, handler cần window)
            AutoPlaceElementFrBlockCADRequestHandler handler = null;
            PlaceElementsFromBlocksCadWpfWindow window = null;
            ExternalEvent exEvent = null;

            // Tạo external event
            handler = new AutoPlaceElementFrBlockCADRequestHandler(viewModel);
            exEvent = ExternalEvent.Create(handler);

            // Sau đó gán lại window cho handler
            window = new PlaceElementsFromBlocksCadWpfWindow(exEvent, handler, viewModel);

            // Show window
            window.ShowDialog();

            return Result.Succeeded;
        }
    }
}
