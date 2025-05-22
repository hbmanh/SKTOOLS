using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
{
    public class AutoPlaceElementFrBlockCADRequestHandler : IExternalEventHandler
    {
        private AutoPlaceElementFrBlockCADViewModel ViewModel;

        public AutoPlaceElementFrBlockCADRequestHandler(AutoPlaceElementFrBlockCADViewModel viewModel)
        {
            ViewModel = viewModel;
        }

        private AutoPlaceElementFrBlockCADRequest m_Request = new AutoPlaceElementFrBlockCADRequest();

        public AutoPlaceElementFrBlockCADRequest Request => m_Request;

        public void Execute(UIApplication uiapp)
        {
            try
            {
                switch (Request.Take())
                {
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

        public string GetName() => "Place Elements from CAD Blocks";

        private void PlaceElements(UIApplication uiapp, AutoPlaceElementFrBlockCADViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var blockMappings = viewModel.BlockMappings;
            var level = viewModel.Level;

            Dictionary<string, int> blockInstanceCounts = new Dictionary<string, int>();

            // Loại trùng tuyệt đối toàn bộ
            HashSet<string> usedPositions = new HashSet<string>();

            using (Transaction trans = new Transaction(doc, "Place Elements from CAD Blocks"))
            {
                trans.Start();
                foreach (var blockMapping in blockMappings)
                {
                    if (!blockMapping.IsEnabled)
                        continue;

                    var blockGroups = blockMapping.Blocks;
                    List<ElementId> createdInstanceIds = new List<ElementId>();

                    foreach (var blockWithLink in blockGroups)
                    {
                        var block = blockWithLink.Block;
                        var cadLink = blockWithLink.CadLink;
                        var category = blockMapping.SelectedCategoryMapping;
                        var selectedType = blockMapping.SelectedTypeSymbolMapping;
                        double offset = blockMapping.Offset / 304.8;

                        if (selectedType == null) continue;
                        if (!selectedType.IsActive)
                        {
                            selectedType.Activate();
                            doc.Regenerate();
                        }

                        Transform importTransform = cadLink.GetTotalTransform();
                        Transform blockTransform = block.Transform;

                        XYZ localOrigin = blockTransform.Origin;
                        XYZ worldOrigin = importTransform.OfPoint(localOrigin);

                        XYZ placementPosition;
                        if (category != null && category.Id.IntegerValue == (int)BuiltInCategory.OST_DetailComponents)
                        {
                            placementPosition = worldOrigin;
                        }
                        else
                        {
                            placementPosition = new XYZ(worldOrigin.X, worldOrigin.Y, worldOrigin.Z + offset);
                        }

                        string posKey = $"{Math.Round(placementPosition.X, 4)},{Math.Round(placementPosition.Y, 4)},{Math.Round(placementPosition.Z, 4)}";
                        if (usedPositions.Contains(posKey))
                            continue;
                        usedPositions.Add(posKey);

                        XYZ globalBasisX = importTransform.OfVector(blockTransform.BasisX);
                        XYZ baseDir = new XYZ(1, 0, 0);

                        double blockRotation = baseDir.AngleTo(globalBasisX);
                        double sign = (XYZ.BasisZ.CrossProduct(baseDir)).DotProduct(globalBasisX) < 0 ? -1 : 1;
                        blockRotation *= sign;

                        FamilyInstance familyInstance = doc.Create.NewFamilyInstance(placementPosition, selectedType, level, StructuralType.NonStructural);
                        ElementTransformUtils.RotateElement(doc, familyInstance.Id, Line.CreateBound(placementPosition, placementPosition + XYZ.BasisZ), blockRotation);

                        createdInstanceIds.Add(familyInstance.Id);

                        if (!blockInstanceCounts.ContainsKey(blockMapping.DisplayBlockName))
                            blockInstanceCounts[blockMapping.DisplayBlockName] = 0;
                        blockInstanceCounts[blockMapping.DisplayBlockName]++;
                    }

                    // Group theo block name nếu có instance, đặt tên cho GroupType
                    if (createdInstanceIds.Count > 0)
                    {
                        Group group = doc.Create.NewGroup(createdInstanceIds);
                        try
                        {
                            if (group.GroupType != null && group.GroupType.CanBeRenamed)
                                group.GroupType.Name = blockMapping.DisplayBlockName;
                        }
                        catch { }
                    }
                }
                trans.Commit();
            }

            ShowResultDialog(blockInstanceCounts);
        }

        private void ShowResultDialog(Dictionary<string, int> blockInstanceCounts)
        {
            TaskDialog dialog = new TaskDialog("Placement Result");
            dialog.MainInstruction = "Placement of Family Instances Complete";
            dialog.MainContent = "The following Family Instances were created for each block:";

            string details = "";
            foreach (var kvp in blockInstanceCounts)
                details += $"{kvp.Key}: {kvp.Value} instances\n";

            dialog.ExpandedContent = details;
            dialog.Show();
        }
    }
}
