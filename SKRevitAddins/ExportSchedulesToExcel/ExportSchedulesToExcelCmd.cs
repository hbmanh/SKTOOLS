using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.ExportSchedulesToExcel
{
    [Transaction(TransactionMode.Manual)]
    public class ExportSchedulesToExcelCmd : IExternalCommand
    {
        private static ExportSchedulesToExcelWpfWindow _modelessWindow;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            try
            {
                // Nếu cửa sổ đã mở, chỉ kích hoạt lại
                if (_modelessWindow != null && _modelessWindow.IsVisible)
                {
                    _modelessWindow.Activate();
                    return Result.Succeeded;
                }

                // Tạo ViewModel, Request và Handler
                var vm = new ExportSchedulesToExcelViewModel(uiApp);
                var request = new ExportSchedulesToExcelRequest();
                var handler = new ExportSchedulesToExcelRequestHandler(vm, request);

                // Tạo ExternalEvent để giao tiếp với API Revit
                var exEvent = ExternalEvent.Create(handler);

                // Tạo và hiển thị cửa sổ modeless
                _modelessWindow = new ExportSchedulesToExcelWpfWindow(exEvent, handler, vm);
                _modelessWindow.Show();

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
