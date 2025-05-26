using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
{
    public class AutoPlaceElementFrBlockCADRequestHandler : IExternalEventHandler
    {
        public AutoPlaceElementFrBlockCADViewModel ViewModel { get; set; }
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

            int total = blockMappings.Sum(bm => bm.Blocks.Count);
            int count = 0;

            using (Transaction trans = new Transaction(doc, "Place Elements from CAD Blocks"))
            {
                trans.Start();

                foreach (var blockMapping in blockMappings)
                {
                    if (!blockMapping.IsEnabled) continue;

                    var createdInstanceIds = new List<ElementId>();
                    int duplicateCount = 0, failedCount = 0;

                    blockMapping.FailureNote = string.Empty;

                    // CHỐNG DUPLICATE: lấy hết instance của family type này trong model
                    var existingPositions = new HashSet<PositionKey>();
                    var selectedType = blockMapping.SelectedTypeSymbolMapping;
                    var category = blockMapping.SelectedCategoryMapping;

                    if (selectedType != null && category != null)
                    {
                        var allInstances = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilyInstance))
                            .OfCategoryId(category.Id)
                            .WhereElementIsNotElementType()
                            .Cast<FamilyInstance>()
                            .Where(fi => fi.Symbol.Id == selectedType.Id && fi.LevelId == level.Id);

                        foreach (var fi in allInstances)
                        {
                            var loc = fi.Location as LocationPoint;
                            if (loc == null) continue;
                            var posKey = new PositionKey(loc.Point);
                            existingPositions.Add(posKey);
                        }
                    }

                    var usedPositions = new HashSet<PositionKey>(existingPositions);
                    var activatedSymbols = new HashSet<ElementId>();

                    foreach (var blockWithLink in blockMapping.Blocks.Where(b => b.IsEnabled))
                    {
                        var block = blockWithLink.Block;
                        var cadLink = blockWithLink.CadLink;
                        selectedType = blockMapping.SelectedTypeSymbolMapping;
                        category = blockMapping.SelectedCategoryMapping;

                        Transform importTransform = cadLink.GetTotalTransform();
                        Transform blockTransform = block.Transform;
                        XYZ localOrigin = blockTransform.Origin;
                        XYZ worldOrigin = importTransform.OfPoint(localOrigin);
                        double offset = blockMapping.Offset / 304.8;
                        XYZ placementPosition = category.Id.IntegerValue == (int)BuiltInCategory.OST_DetailComponents
                            ? worldOrigin
                            : new XYZ(worldOrigin.X, worldOrigin.Y, worldOrigin.Z + offset);

                        var posKey = new PositionKey(placementPosition);

                        double angle = new XYZ(1, 0, 0).AngleTo(blockTransform.BasisX);

                        if (selectedType == null || cadLink == null)
                        {
                            failedCount++;
                            continue;
                        }

                        if (!activatedSymbols.Contains(selectedType.Id))
                        {
                            if (!selectedType.IsActive)
                                selectedType.Activate();
                            activatedSymbols.Add(selectedType.Id);
                        }

                        if (usedPositions.Contains(posKey))
                        {
                            duplicateCount++;
                            continue;
                        }
                        usedPositions.Add(posKey);

                        try
                        {
                            var instance = doc.Create.NewFamilyInstance(placementPosition, selectedType, level, StructuralType.NonStructural);

                            ElementTransformUtils.RotateElement(doc, instance.Id,
                                Line.CreateBound(placementPosition, placementPosition + XYZ.BasisZ), angle);

                            createdInstanceIds.Add(instance.Id);
                            count++;
                            viewModel.UpdateStatus?.Invoke($"Đã đặt {count}/{total} block...");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"  ERROR: {ex.Message}");
                            failedCount++;
                        }
                    }

                    blockMapping.PlacedCount = createdInstanceIds.Count;
                    blockMapping.HasPlacementRun = true;
                    blockMapping.OnPropertyChanged(nameof(blockMapping.RowStatusColor));

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

                    if (duplicateCount > 0)
                        blockMapping.FailureNote = $"Có {duplicateCount} block trùng lặp vị trí.";
                    else if (failedCount > 0)
                        blockMapping.FailureNote = $"Có {failedCount} block thiếu Family Type/CAD Link.";
                    else if (createdInstanceIds.Count < blockMapping.Blocks.Count(b => b.IsEnabled))
                        blockMapping.FailureNote = "Không đặt đủ instance (lý do khác).";
                    else
                        blockMapping.FailureNote = string.Empty;
                }

                doc.Regenerate();
                viewModel.OnPropertyChanged(nameof(viewModel.BlockMappings));
                trans.Commit();
            }
        }
    }
}
