//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.UI;

//namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
//{
//    [Transaction(TransactionMode.Manual)]
//    public class AutoPlaceElementFrBlockCADCmd : IExternalCommand
//    {
//        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//        {
//            UIApplication uiapp = commandData.Application;
//            UIDocument uidoc = uiapp.ActiveUIDocument;

//            var handler = new AutoPlaceElementFrBlockCADRequestHandler();
//            var exEvent = ExternalEvent.Create(handler);

//            var viewModel = new AutoPlaceElementFrBlockCADViewModel(uiapp, requestId =>
//            {
//                handler.Request.Make(requestId);
//                exEvent.Raise();
//            });

//            handler.ViewModel = viewModel;

//            var window = new PlaceElementsFromBlocksCadWpfWindow(exEvent, handler, viewModel);
//            window.Show();

//            return Result.Succeeded;
//        }
//    }
//}
