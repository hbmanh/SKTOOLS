using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.PointCloudAddins.ScanAndModel
{
    [Transaction(TransactionMode.Manual)]
    public class ScanAndModelCmd : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            try
            {
                // 1) Tạo ViewModel
                var viewModel = new ScanAndModelViewModel(uiApp);

                // 2) Tạo Request + RequestHandler
                var request = new ScanAndModelRequest();
                var handler = new ScanAndModelRequestHandler(viewModel, request);

                // 3) Tạo ExternalEvent
                var exEvent = ExternalEvent.Create(handler);

                // 4) Mở cửa sổ chính
                var mainWindow = new ScanAndModelWpfWindow(exEvent, handler, viewModel);
                mainWindow.ShowDialog();

                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
