using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using SKRevitAddins.Commands.ExportSchedulesToExcel;
using SKRevitAddins.ViewModel; // Nếu bạn đặt ViewModel chung
using SKRevitAddins.Forms;

namespace SKRevitAddins.Commands.ExportSchedulesToExcel
{
    [Transaction(TransactionMode.Manual)]
    public class ExportSchedulesToExcelCmd : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            // Khởi tạo ViewModel
            ExportSchedulesToExcelViewModel viewModel = new ExportSchedulesToExcelViewModel(uiApp);
            try
            {
                // Gọi hàm hiển thị cửa sổ WPF được định nghĩa trong App
                App.thisApp.ShowExportSchedulesToExcelViewModel(uiApp, viewModel);
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
