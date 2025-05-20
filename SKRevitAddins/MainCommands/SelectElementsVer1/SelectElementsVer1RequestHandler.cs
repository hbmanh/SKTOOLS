using System;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Document = Autodesk.Revit.DB.Document;
using View = Autodesk.Revit.DB.View;

namespace SKRevitAddins.SelectElementsVer1
{
    public class SelectElementsVer1RequestHandler : IExternalEventHandler
    {
        private SelectElementsVer1ViewModel ViewModel;

        public SelectElementsVer1RequestHandler(SelectElementsVer1ViewModel viewModel)
        {
            ViewModel = viewModel;
        }

        private SelectElementsVer1Request m_Request = new SelectElementsVer1Request();

        public SelectElementsVer1Request Request
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

        public void NumberingElements(UIApplication uiapp, SelectElementsVer1ViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var selFilterBy = viewModel.SelFilterBy;
            var elementsToChange = viewModel.EleToPreview;
            var selParameter= viewModel.SelParameter;
            var isUpToDown = viewModel.IsUpToDown;
            var isLeftToRight = viewModel.IsLeftToRight;
            var isRightToLeft = viewModel.IsRightToLeft;
            var isDownToUp = viewModel.IsDownToUp;
            var beginWith = viewModel.BeginsWith;
            var keywords = viewModel.Keywords;
            var keyTarget = viewModel.KeyTarget;

            if (isUpToDown)
            {
                elementsToChange = elementsToChange.OrderByDescending(e => GetBoundingBox(e, doc.ActiveView).Min.Y).ToList();
            }

            if (isLeftToRight)
            {
                elementsToChange = elementsToChange.OrderBy(e => GetBoundingBox(e, doc.ActiveView).Min.X).ToList();
            }

            if (isRightToLeft)
            {
                elementsToChange = elementsToChange.OrderByDescending(e => GetBoundingBox(e, doc.ActiveView).Min.X).ToList();
            }

            if (isDownToUp)
            {
                elementsToChange = elementsToChange.OrderBy(e => GetBoundingBox(e, doc.ActiveView).Min.Y).ToList();
            }

            ObservableCollection<Element> sortedElements = new ObservableCollection<Element>(elementsToChange);

            using (Transaction tx = new Transaction(doc, "Modify Param"))
            {
                tx.Start();
                if (selFilterBy == "Numbering")
                {
                    int count = 1;
                    foreach (var element in elementsToChange)
                    {
                        string number = beginWith + count.ToString();
                        Parameter param = element.LookupParameter(selParameter.Definition.Name);
                        if (param != null && param.StorageType == StorageType.String)
                        {
                            param.Set(number);
                        }
                        count++;
                    }
                }
                if (selFilterBy == "Replace symbol")
                {
                    foreach (var element in elementsToChange)
                    {
                        string currentValue = element.LookupParameter(selParameter.Definition.Name)?.AsString();
                        if (currentValue != null && currentValue.Contains(keywords))
                        {
                            string newNumber = currentValue.Replace(keywords, keyTarget);
                            Parameter param = element.LookupParameter(selParameter.Definition.Name);
                            if (param != null && param.StorageType == StorageType.String)
                            {
                                param.Set(newNumber);
                            }
                        }
                    }
                }
                
                tx.Commit();
            }
        }

        private BoundingBoxXYZ GetBoundingBox(Element element, View view)
        {
            BoundingBoxXYZ boundingBox = element.get_BoundingBox(view);
            return boundingBox;
        }

        #endregion

    }
}
