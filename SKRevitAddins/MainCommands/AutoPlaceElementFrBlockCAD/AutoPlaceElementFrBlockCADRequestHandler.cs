using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;

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

            var blockInstanceCounts = new Dictionary<string, int>();
            var usedPositions = new HashSet<PositionKey>();
            var activatedSymbols = new HashSet<ElementId>();

            int total = blockMappings.Sum(bm => bm.Blocks.Count);
            int count = 0;

            using (Transaction trans = new Transaction(doc, "Place Elements from CAD Blocks"))
            {
                trans.Start();

                foreach (var blockMapping in blockMappings)
                {
                    if (!blockMapping.IsEnabled) continue;

                    var createdInstanceIds = new List<ElementId>();

                    foreach (var blockWithLink in blockMapping.Blocks)
                    {
                        var block = blockWithLink.Block;
                        var cadLink = blockWithLink.CadLink;
                        var selectedType = blockMapping.SelectedTypeSymbolMapping;
                        var category = blockMapping.SelectedCategoryMapping;

                        if (selectedType == null || cadLink == null) continue;

                        if (!activatedSymbols.Contains(selectedType.Id))
                        {
                            if (!selectedType.IsActive)
                                selectedType.Activate();

                            activatedSymbols.Add(selectedType.Id);
                        }

                        Transform importTransform = cadLink.GetTotalTransform();
                        Transform blockTransform = block.Transform;

                        XYZ localOrigin = blockTransform.Origin;
                        XYZ worldOrigin = importTransform.OfPoint(localOrigin);
                        double offset = blockMapping.Offset / 304.8;

                        XYZ placementPosition = category.Id.IntegerValue == (int)BuiltInCategory.OST_DetailComponents
                            ? worldOrigin
                            : new XYZ(worldOrigin.X, worldOrigin.Y, worldOrigin.Z + offset);

                        var posKey = new PositionKey(placementPosition);
                        if (usedPositions.Contains(posKey)) continue;
                        usedPositions.Add(posKey);

                        XYZ baseDir = new XYZ(1, 0, 0);
                        XYZ globalBasisX = importTransform.OfVector(blockTransform.BasisX);

                        double angle = baseDir.AngleTo(globalBasisX);
                        double sign = XYZ.BasisZ.CrossProduct(baseDir).DotProduct(globalBasisX) < 0 ? -1 : 1;
                        double blockRotation = angle * sign;

                        var instance = doc.Create.NewFamilyInstance(placementPosition, selectedType, level, StructuralType.NonStructural);
                        ElementTransformUtils.RotateElement(doc, instance.Id, Line.CreateBound(placementPosition, placementPosition + XYZ.BasisZ), blockRotation);

                        createdInstanceIds.Add(instance.Id);

                        if (!blockInstanceCounts.ContainsKey(blockMapping.DisplayBlockName))
                            blockInstanceCounts[blockMapping.DisplayBlockName] = 0;

                        blockInstanceCounts[blockMapping.DisplayBlockName]++;
                        count++;
                        viewModel.UpdateStatus?.Invoke($"Đã đặt {count}/{total} block...");
                    }

                    if (createdInstanceIds.Count > 0)
                    {
                        var group = doc.Create.NewGroup(createdInstanceIds);
                        try
                        {
                            if (group.GroupType?.CanBeRenamed == true)
                                group.GroupType.Name = blockMapping.DisplayBlockName;
                        }
                        catch { }
                    }
                }

                doc.Regenerate();
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
