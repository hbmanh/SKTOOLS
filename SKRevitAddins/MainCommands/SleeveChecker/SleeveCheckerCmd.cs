using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.SleeveChecker
{
    [Transaction(TransactionMode.Manual)]
    public class SleeveCheckerCmd : IExternalCommand
    {
        private static SleeveCheckerWindow _window = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            var viewModel = new SleeveCheckerViewModel(uiapp);
            var handler = new CheckerRequestHandler(viewModel);
            var exEvent = ExternalEvent.Create(handler);
            ShowWindow(exEvent, handler, viewModel);
            return Result.Succeeded;
        }

        private void ShowWindow(ExternalEvent exEvent, CheckerRequestHandler handler, SleeveCheckerViewModel viewModel)
        {
            if (_window == null || !_window.IsVisible)
            {
                _window = new SleeveCheckerWindow(exEvent, handler, viewModel);
                _window.Show();
            }
            else
            {
                _window.Activate();
            }
        }
    }
}