using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace SKRevitAddins.LayoutsToDWG
{
    [Transaction(TransactionMode.Manual)]
    public class LayoutsToDWGCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
            ref string message,
            Autodesk.Revit.DB.ElementSet elements)
        {
            var win = new LayoutsToDWGWindow(commandData.Application.ActiveUIDocument);
            win.ShowDialog();
            return Result.Succeeded;
        }
    }
}