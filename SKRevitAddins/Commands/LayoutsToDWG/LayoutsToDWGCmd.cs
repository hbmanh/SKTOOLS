using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    [Transaction(TransactionMode.Manual)]
    public class LayoutsToDWGCmd : IExternalCommand
    {
        private static LayoutsToDWGWindow _window;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                if (_window != null && _window.IsVisible)
                {
                    _window.Activate();
                    return Result.Succeeded;
                }

                var vm = new LayoutsToDWGViewModel(commandData.Application);
                var handler = new LayoutsToDWGRequestHandler(vm);
                var ev = ExternalEvent.Create(handler);

                _window = new LayoutsToDWGWindow(ev, handler, vm);
                _window.Show();

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
