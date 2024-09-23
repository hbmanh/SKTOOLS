using System;
using Autodesk.Revit.UI;
using SKRevitAddins.ViewModel;
using Document = Autodesk.Revit.DB.Document;

namespace SKRevitAddins.Commands.DeleteTypeOfTextNotesDontUse
{
    public class DeleteTypeOfTextNotesDontUseRequestHandler : IExternalEventHandler
    {
        public Document Doc;
        private DeleteTypeOfTextNotesDontUseViewModel ViewModel;

        public DeleteTypeOfTextNotesDontUseRequestHandler(DeleteTypeOfTextNotesDontUseViewModel viewModel)
        {
            ViewModel = viewModel;
        }

        private DeleteTypeOfTextNotesDontUseRequest m_Request = new DeleteTypeOfTextNotesDontUseRequest();

        public DeleteTypeOfTextNotesDontUseRequest Request
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
                        DelteTypeOfTextNotesDontUse(uiapp, ViewModel);
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

        #region Delete Type of Text Notes Dont Use

        public void DelteTypeOfTextNotesDontUse(UIApplication uiapp, DeleteTypeOfTextNotesDontUseViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
        }

        #endregion
    }
}
