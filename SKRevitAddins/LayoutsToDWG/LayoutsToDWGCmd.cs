using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.LayoutsToDWG
{
    [Transaction(TransactionMode.Manual)]
    public class LayoutsToDWGCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData data,
            ref string message,
            ElementSet elements)
        {
            var handler = new LayoutsToDWGRequestHandler();
            var exEvent = ExternalEvent.Create(handler);
            var vm = new LayoutsToDWGViewModel(
                data.Application.ActiveUIDocument,
                exEvent,
                handler);

            var win = new LayoutsToDWGWindow(exEvent, handler, vm);
            win.ShowDialog();          // modal
            return Result.Succeeded;
        }
    }
}