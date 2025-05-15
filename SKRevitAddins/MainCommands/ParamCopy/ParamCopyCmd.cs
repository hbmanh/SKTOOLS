//using Autodesk.Revit.ApplicationServices;
//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.UI;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Autodesk.Revit.UI.Selection;

//namespace ParamCopy
//{
//    [Transaction(TransactionMode.Manual)]
//    public class ParamCopyCmd : IExternalCommand
//    {
//        ParamCopyViewModel viewModel;
//        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//        {
//            UIApplication uiapp = commandData.Application;
//            UIDocument uidoc = uiapp.ActiveUIDocument;
//            Application app = uiapp.Application;
//            Document doc = uidoc.Document;

//            viewModel = new ParamCopyViewModel(uiapp);

//            try
//            {
//                App.thisApp.ShowParamCopyWindow(commandData.Application, viewModel);
//                return Result.Succeeded;
//            }
//            catch (Exception ex)
//            {
//                message = ex.Message;
//                return Result.Failed;
//            }
//        }
//    }
//}
