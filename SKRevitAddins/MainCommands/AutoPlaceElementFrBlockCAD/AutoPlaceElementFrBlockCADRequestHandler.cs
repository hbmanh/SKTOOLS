using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System.Diagnostics;

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
                    int duplicateCount = 0, failedCount = 0;

                    blockMapping.FailureNote = string.Empty;

                    foreach (var blockWithLink in blockMapping.Blocks)
                    {
                        var block = blockWithLink.Block;
                        var cadLink = blockWithLink.CadLink;
                        var selectedType = blockMapping.SelectedTypeSymbolMapping;
                        var category = blockMapping.SelectedCategoryMapping;

                        Transform importTransform = cadLink.GetTotalTransform();
                        Transform blockTransform = block.Transform;
                        XYZ localOrigin = blockTransform.Origin;
                        XYZ worldOrigin = importTransform.OfPoint(localOrigin);

                        double offset = blockMapping.Offset / 304.8;

                        XYZ placementPosition = category.Id.IntegerValue == (int)BuiltInCategory.OST_DetailComponents
                            ? worldOrigin
                            : new XYZ(worldOrigin.X, worldOrigin.Y, worldOrigin.Z + offset);

                        var posKey = new PositionKey(placementPosition, blockTransform);
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
                    blockMapping.OnPropertyChanged(nameof(blockMapping.RowStatusColor)); // ✅ cập nhật màu

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
                    else if (createdInstanceIds.Count < blockMapping.Blocks.Count)
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

    public class PositionKey
    {
        public readonly double x, y, z;
        public readonly double scaleX, scaleY, scaleZ;
        public readonly double rotation; // Thêm rotation để phân biệt

        private const double POSITION_TOLERANCE = 0.003; // ~1mm
        private const double SCALE_TOLERANCE = 0.001;    // Tolerance cho scale
        private const double ROTATION_TOLERANCE = 0.017; // ~1 degree

        public PositionKey(XYZ position, Transform blockTransform)
        {
            // Position
            x = Math.Round(position.X / POSITION_TOLERANCE) * POSITION_TOLERANCE;
            y = Math.Round(position.Y / POSITION_TOLERANCE) * POSITION_TOLERANCE;
            z = Math.Round(position.Z / POSITION_TOLERANCE) * POSITION_TOLERANCE;

            // Scale từ Transform matrix
            scaleX = Math.Round(blockTransform.BasisX.GetLength() / SCALE_TOLERANCE) * SCALE_TOLERANCE;
            scaleY = Math.Round(blockTransform.BasisY.GetLength() / SCALE_TOLERANCE) * SCALE_TOLERANCE;
            scaleZ = Math.Round(blockTransform.BasisZ.GetLength() / SCALE_TOLERANCE) * SCALE_TOLERANCE;

            // Rotation (angle của BasisX với X-axis)
            XYZ baseDir = new XYZ(1, 0, 0);
            double angle = baseDir.AngleTo(blockTransform.BasisX);
            rotation = Math.Round(angle / ROTATION_TOLERANCE) * ROTATION_TOLERANCE;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is PositionKey key)) return false;

            return Math.Abs(key.x - x) < POSITION_TOLERANCE &&
                   Math.Abs(key.y - y) < POSITION_TOLERANCE &&
                   Math.Abs(key.z - z) < POSITION_TOLERANCE &&
                   Math.Abs(key.scaleX - scaleX) < SCALE_TOLERANCE &&
                   Math.Abs(key.scaleY - scaleY) < SCALE_TOLERANCE &&
                   Math.Abs(key.scaleZ - scaleZ) < SCALE_TOLERANCE &&
                   Math.Abs(key.rotation - rotation) < ROTATION_TOLERANCE;
        }

        public override int GetHashCode()
        {
            int gridX = (int)Math.Round(x / POSITION_TOLERANCE);
            int gridY = (int)Math.Round(y / POSITION_TOLERANCE);
            int gridZ = (int)Math.Round(z / POSITION_TOLERANCE);
            int gridSX = (int)Math.Round(scaleX / SCALE_TOLERANCE);
            int gridSY = (int)Math.Round(scaleY / SCALE_TOLERANCE);
            int gridSZ = (int)Math.Round(scaleZ / SCALE_TOLERANCE);
            int gridR = (int)Math.Round(rotation / ROTATION_TOLERANCE);

            return HashCode.Combine(gridX, gridY, gridZ, gridSX, gridSY, gridSZ, gridR);
        }

        public override string ToString()
        {
            return $"Pos:({x:F3},{y:F3},{z:F3}) Scale:({scaleX:F3},{scaleY:F3},{scaleZ:F3}) Rot:{rotation:F3}";
        }
    }
}
