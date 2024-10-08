using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;
using SKRevitAddins.ViewModel;
using Document = Autodesk.Revit.DB.Document;

namespace SKRevitAddins.Commands.PermissibleRangeFrame
{
    public class PermissibleRangeFrameRequestHandler : IExternalEventHandler
    {
        private PermissibleRangeFrameViewModel ViewModel;


        public PermissibleRangeFrameRequestHandler(PermissibleRangeFrameViewModel viewModel)
        {
            ViewModel = viewModel;
        }

        private PermissibleRangeFrameRequest m_Request = new PermissibleRangeFrameRequest();

        public PermissibleRangeFrameRequest Request
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
                        AutoCreatePileFromCad(uiapp, ViewModel);
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
            return "";
        }

        #region Auto Create Pile From Cad
        public void AutoCreatePileFromCad(UIApplication uiapp, PermissibleRangeFrameViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            double offset = viewModel.Offset;
            Level selectedLevel = viewModel.SelectedLevel;
            var selectedPileType = viewModel.SelectedPileType;
            List<Arc> allPileData = CadUtils.GetArcsHaveName(viewModel.SelectedCadLink, viewModel.SelectedLayer);
            if (allPileData.Count <= 0) return;
            List<FamilyInstance> newPiles = new List<FamilyInstance>();
            viewModel.Offset = 0;
            using (Transaction txCreatePile = new Transaction(doc))
            {
                txCreatePile.Start("Create Pile");
                foreach (Arc aLine in allPileData)
                {
                    var center = aLine.Center;
                    var newPile = doc.Create.NewFamilyInstance(center, selectedPileType, selectedLevel, Autodesk.Revit.DB.Structure.StructuralType.Footing);

                    //Set level
                    newPile.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).Set(selectedLevel.Id);
                    newPile.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(Utils.UnitUtils.MmToFeet(selectedLevel.Elevation + offset));
                    newPiles.Add(newPile);
                }
                txCreatePile.Commit();
            }
        }
        #endregion
    }
}
