using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.PermissibleRangeFrame
{
    [Transaction(TransactionMode.Manual)]
    public class PermissibleRangeFrameCmd : IExternalCommand
    {
        private static PermissibleRangeFrameWpfWindow _window = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            var viewModel = new PermissibleRangeFrameViewModel(uiapp);
            ShowPermissibleRangeFrameWindow(uiapp, viewModel);
            return Result.Succeeded;
        }

        private void ShowPermissibleRangeFrameWindow(UIApplication uiapp, PermissibleRangeFrameViewModel viewModel)
        {
            if (_window == null || !_window.IsVisible)
            {
                var handler = new PermissibleRangeFrameRequestHandler(viewModel);
                var exEvent = ExternalEvent.Create(handler);
                _window = new PermissibleRangeFrameWpfWindow(exEvent, handler, viewModel);
                _window.Show();
            }
            else
            {
                _window.Activate();
            }
        }
    }
}