using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKRevitAddins.ViewModel;
using Document = Autodesk.Revit.DB.Document;

namespace SKRevitAddins.Commands.FindDWGNotUsedAndDel
{
    public class FindDWGNotUsedAndDelRequestHandler : IExternalEventHandler
    {
        private FindDWGNotUsedAndDelViewModel ViewModel;

        public FindDWGNotUsedAndDelRequestHandler(FindDWGNotUsedAndDelViewModel viewModel)
        {
            ViewModel = viewModel;
        }

        private FindDWGNotUsedAndDelRequest m_Request = new FindDWGNotUsedAndDelRequest();

        public FindDWGNotUsedAndDelRequest Request
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
                        SelectElements(uiapp, ViewModel);
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

        public void SelectElements(UIApplication uiapp, FindDWGNotUsedAndDelViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            
        }

        #endregion

        private DefinitionGroup GetOrCreateGroup(DefinitionFile sharedParamFile, string groupName)
        {
            foreach (DefinitionGroup group in sharedParamFile.Groups)
            {
                if (group.Name == groupName)
                {
                    return group;
                }
            }
            return sharedParamFile.Groups.Create(groupName);
        }
       

    }
}
