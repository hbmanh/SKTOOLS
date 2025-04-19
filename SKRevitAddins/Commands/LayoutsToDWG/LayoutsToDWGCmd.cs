// SKRevitAddins.Commands.LayoutsToDWG.LayoutsToDWGCmd.cs
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    [Transaction(TransactionMode.Manual)]
    public class LayoutsToDWGCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData cData, ref string message, ElementSet elements)
        {
            var win = new LayoutsToDWGWindow(cData.Application.ActiveUIDocument);
            win.ShowDialog();
            return Result.Succeeded;
        }
    }
}