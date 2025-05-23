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

                        // === DEBUG: LOG THÔNG TIN CHI TIẾT ===
                        System.Text.StringBuilder log = new System.Text.StringBuilder();
                        log.AppendLine("=========================================");
                        log.AppendLine($"Block Name: {blockMapping.DisplayBlockName}");
                        log.AppendLine($"Block InstanceId: {block.Id}");

                        Transform importTransform = cadLink.GetTotalTransform();
                        Transform blockTransform = block.Transform;
                        XYZ localOrigin = blockTransform.Origin;
                        XYZ worldOrigin = importTransform.OfPoint(localOrigin);

                        log.AppendLine($"  Local Origin: ({localOrigin.X:F4}, {localOrigin.Y:F4}, {localOrigin.Z:F4})");
                        log.AppendLine($"  World Origin: ({worldOrigin.X:F4}, {worldOrigin.Y:F4}, {worldOrigin.Z:F4})");
                        log.AppendLine($"  BasisX: ({blockTransform.BasisX.X:F4}, {blockTransform.BasisX.Y:F4}, {blockTransform.BasisX.Z:F4}), Length: {blockTransform.BasisX.GetLength():F4}");
                        log.AppendLine($"  BasisY: ({blockTransform.BasisY.X:F4}, {blockTransform.BasisY.Y:F4}, {blockTransform.BasisY.Z:F4}), Length: {blockTransform.BasisY.GetLength():F4}");
                        log.AppendLine($"  BasisZ: ({blockTransform.BasisZ.X:F4}, {blockTransform.BasisZ.Y:F4}, {blockTransform.BasisZ.Z:F4}), Length: {blockTransform.BasisZ.GetLength():F4}");

                        double offset = blockMapping.Offset / 304.8;

                        XYZ placementPosition = category.Id.IntegerValue == (int)BuiltInCategory.OST_DetailComponents
                            ? worldOrigin
                            : new XYZ(worldOrigin.X, worldOrigin.Y, worldOrigin.Z + offset);

                        // DEBUG: PositionKey
                        var posKey = new PositionKey(placementPosition, blockTransform);
                        log.AppendLine($"  PositionKey: {posKey}");

                        // Rotation angle (so với trục X gốc)
                        XYZ baseDir = new XYZ(1, 0, 0);
                        double angle = baseDir.AngleTo(blockTransform.BasisX);
                        double angleDeg = angle * 180.0 / Math.PI;
                        log.AppendLine($"  Rotation angle to X axis: {angleDeg:F2} deg");

                        if (usedPositions.Contains(posKey))
                        {
                            log.AppendLine("  >>> DUPLICATE detected, block will NOT be placed.");
                        }
                        else
                        {
                            log.AppendLine("  >>> Unique position, block will be placed.");
                        }

                        // In ra log (cả Output Window lẫn file, nếu cần)
                        Debug.WriteLine(log.ToString());
                        System.IO.File.AppendAllText(@"C:\Temp\block_debug_log.txt", log.ToString());

                        // GHI FILE (nếu muốn), nhớ tạo sẵn thư mục C:\Temp hoặc đổi đường dẫn
                        // System.IO.File.AppendAllText(@"C:\Temp\block_debug_log.txt", log.ToString());

                        // --- PHẦN CODE GỐC ---
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
                            // ... code đặt FamilyInstance như cũ ...
                            var instance = doc.Create.NewFamilyInstance(placementPosition, selectedType, level, StructuralType.NonStructural);

                            double scaleFactorX = blockTransform.BasisX.GetLength();
                            double scaleFactorY = blockTransform.BasisY.GetLength();
                            double scaleFactorZ = blockTransform.BasisZ.GetLength();

                            // Apply rotation
                            ElementTransformUtils.RotateElement(doc, instance.Id,
                                Line.CreateBound(placementPosition, placementPosition + XYZ.BasisZ), angle);

                            // Nếu scale khác 1.0, log cảnh báo (Revit ko scale được family thường)
                            if (Math.Abs(scaleFactorX - 1.0) > 0.01 || Math.Abs(scaleFactorY - 1.0) > 0.01)
                            {
                                Debug.WriteLine($"  *** WARNING: Non-uniform scale detected but not applied to Revit instance ***");
                                // TODO: Implement scaling if needed
                            }

                            createdInstanceIds.Add(instance.Id);
                            count++;
                            viewModel.UpdateStatus?.Invoke($"Đã đặt {count}/{total} block...");

                            Debug.WriteLine($"  *** SUCCESS: Created instance {instance.Id} ***");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"  *** ERROR creating instance: {ex.Message} ***");
                            failedCount++;
                        }
                    }

                    blockMapping.PlacedCount = createdInstanceIds.Count;
                    blockMapping.HasPlacementRun = true;

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

                    // Update failure notes
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
