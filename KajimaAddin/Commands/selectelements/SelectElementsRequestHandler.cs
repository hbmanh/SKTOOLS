using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKToolsAddins.ViewModel;
using Document = Autodesk.Revit.DB.Document;
using View = Autodesk.Revit.DB.View;

namespace SKToolsAddins.Commands.SelectElements
{
    public class SelectElementsRequestHandler : IExternalEventHandler
    {
        private SelectElementsViewModel ViewModel;

        public SelectElementsRequestHandler(SelectElementsViewModel viewModel)
        {
            ViewModel = viewModel;
        }

        private SelectElementsRequest m_Request = new SelectElementsRequest();

        public SelectElementsRequest Request
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
                        NumberingElements(uiapp, ViewModel);
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
            return "パラメータの値でフィルターし、選択するツール";
        }

        #region NumberingElements

        public void NumberingElements(UIApplication uiapp, SelectElementsViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            
        }

        private BoundingBoxXYZ GetBoundingBox(Element element, View view)
        {
            BoundingBoxXYZ boundingBox = element.get_BoundingBox(view);
            return boundingBox;
        }

        #endregion

    }
}
