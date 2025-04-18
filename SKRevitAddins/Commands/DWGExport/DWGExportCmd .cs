using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace SKRevitAddins.Commands.DWGExport
{
    [Transaction(TransactionMode.Manual)]
    public class DWGExportCmd : IExternalCommand
    {
        private static DWGExportWpfWindow _window;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (_window != null && _window.IsVisible)
                {
                    _window.Activate();
                    return Result.Succeeded;
                }

                var vm = new DWGExportViewModel(commandData.Application);
                var req = new DWGExportRequest();
                var handler = new DWGExportRequestHandler(vm, req);
                var ev = ExternalEvent.Create(handler);

                _window = new DWGExportWpfWindow(ev, handler, vm);
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