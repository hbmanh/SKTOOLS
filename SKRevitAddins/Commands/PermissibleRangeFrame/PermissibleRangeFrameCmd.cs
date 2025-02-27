﻿using System;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKRevitAddins.ViewModel;

namespace SKRevitAddins.Commands.PermissibleRangeFrame
{
    [Transaction(TransactionMode.Manual)]
    public class PermissibleRangeFrameCmd : IExternalCommand
    {
        PermissibleRangeFrameViewModel viewModel;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            viewModel = new PermissibleRangeFrameViewModel(uiapp);

            try
            {
                App.thisApp.ShowPermissibleRangeFrameViewModel(uiapp, viewModel);
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
