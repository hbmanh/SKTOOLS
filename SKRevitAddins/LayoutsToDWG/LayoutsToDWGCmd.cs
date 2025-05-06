using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using System.Windows;
using System.Windows.Interop;

namespace SKRevitAddins.LayoutsToDWG
{
    [Transaction(TransactionMode.Manual)]
    public class LayoutsToDWGCmd : IExternalCommand
    {
        static LayoutsToDWGWindow _win;

        public Result Execute(ExternalCommandData cd,
            ref string msg,
            Autodesk.Revit.DB.ElementSet elems)
        {
            if (_win != null) { _win.Activate(); return Result.Succeeded; }

            _win = new LayoutsToDWGWindow(cd.Application.ActiveUIDocument);
            new WindowInteropHelper(_win).Owner = cd.Application.MainWindowHandle;
            _win.Closed += (_, __) => _win = null;

            _win.ShowDialog();            // modal => API context còn sống
            return Result.Succeeded;
        }
    }
}