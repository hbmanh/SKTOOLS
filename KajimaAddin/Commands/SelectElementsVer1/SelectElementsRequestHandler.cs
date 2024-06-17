using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Win32;
using SKToolsAddins.ViewModel;
using Binding = Autodesk.Revit.DB.Binding;
using Document = Autodesk.Revit.DB.Document;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;

namespace SKToolsAddins.Commands.SelectElementsVer1
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

        #region SelectElements

        public void SelectElements(UIApplication uiapp, SelectElementsViewModel viewModel)
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
