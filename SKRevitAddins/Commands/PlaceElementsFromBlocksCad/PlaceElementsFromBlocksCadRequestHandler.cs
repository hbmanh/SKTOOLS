using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using SKRevitAddins.ViewModel;
using Document = Autodesk.Revit.DB.Document;

namespace SKRevitAddins.Commands.PlaceElementsFromBlocksCad
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
                    case RequestId.OK:
                        PlaceElements(uiapp, ViewModel);
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        public string GetName()
        {
            return "Place Elements from CAD Blocks";
        }

        #region Place Elements
        private void PlaceElements(UIApplication uiapp, PlaceElementsFromBlocksCadViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var blockMappings = viewModel.BlockMappings;
            var level = viewModel.Level;

            Dictionary<string, int> blockInstanceCounts = new Dictionary<string, int>();

            using (Transaction trans = new Transaction(doc, "Place Elements from CAD Blocks"))
            {
                trans.Start();
                foreach (var blockMapping in blockMappings)
                {
                    var blockGroups = blockMapping.Blocks;
                    foreach (var block in blockGroups)
                    {
                        var category = blockMapping.SelectedCategoryMapping;
                        var selectedType = blockMapping.SelectedTypeSymbolMapping;
                        double offset = blockMapping.Offset / 304.8;

                        var blockPosition = block.Transform.Origin;
                        var blockRotation = block.Transform.BasisX.AngleTo(new XYZ(1, 0, 0));

                        if (selectedType == null) continue;
                        if (!selectedType.IsActive)
                        {
                            selectedType.Activate();
                            doc.Regenerate();
                        }

                        // If category is detail item, do not apply offset
                        XYZ placementPosition;
                        if (category != null && category.Id.IntegerValue == (int)BuiltInCategory.OST_DetailComponents)
                        {
                            placementPosition = blockPosition;
                        }
                        else
                        {
                            placementPosition = new XYZ(blockPosition.X, blockPosition.Y, offset);
                        }

                        FamilyInstance familyInstance = doc.Create.NewFamilyInstance(placementPosition, selectedType, level, StructuralType.NonStructural);

                        // Apply the rotation
                        ElementTransformUtils.RotateElement(doc, familyInstance.Id, Line.CreateBound(placementPosition, placementPosition + XYZ.BasisZ), blockRotation);

                        // Count the created instances for each block
                        if (!blockInstanceCounts.ContainsKey(blockMapping.DisplayBlockName))
                        {
                            blockInstanceCounts[blockMapping.DisplayBlockName] = 0;
                        }
                        blockInstanceCounts[blockMapping.DisplayBlockName]++;
                    }

                }
                trans.Commit();
            }

            // Show the result dialog
            ShowResultDialog(blockInstanceCounts);
        }

        private void ShowResultDialog(Dictionary<string, int> blockInstanceCounts)
        {
            TaskDialog dialog = new TaskDialog("Placement Result");
            dialog.MainInstruction = "Placement of Family Instances Complete";
            dialog.MainContent = "The following Family Instances were created for each block:";

            string details = "";
            foreach (var kvp in blockInstanceCounts)
            {
                details += $"{kvp.Key}: {kvp.Value} instances\n";
            }

            dialog.ExpandedContent = details;
            dialog.Show();
        }
        #endregion
    }
}
