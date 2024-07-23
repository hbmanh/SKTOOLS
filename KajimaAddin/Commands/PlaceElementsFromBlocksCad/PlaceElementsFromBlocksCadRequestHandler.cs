using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKToolsAddins.Commands.PlaceElementsFromBlocksCad;
using SKToolsAddins.ViewModel;
using Document = Autodesk.Revit.DB.Document;

namespace SKToolsAddins.Commands.PlaceElementsFromBlocksCad
{
    public class PlaceElementsFromBlocksCadRequestHandler : IExternalEventHandler
    {
        public Document Doc;
        private PlaceElementsFromBlocksCadViewModel ViewModel;

        public PlaceElementsFromBlocksCadRequestHandler(PlaceElementsFromBlocksCadViewModel viewModel)
        {
            ViewModel = viewModel;
        }

        private PlaceElementsFromBlocksCadRequest m_Request = new PlaceElementsFromBlocksCadRequest();

        public PlaceElementsFromBlocksCadRequest Request
        {
            get { return m_Request; }
        }

        public void Execute(UIApplication uiapp)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        break;
                    case (RequestId.OK):
                        CopyFilterOptionData(uiapp, ViewModel);
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("エラー", ex.Message);
            }
        }

        public string GetName()
        {
            return "フィルター色コピー";
        }

        #region Copy Filter Option Data
        public void CopyFilterOptionData(UIApplication uiapp, PlaceElementsFromBlocksCadViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            
        }
        #endregion
    }
}
