using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RefPointTopoTool
{
    [Transaction(TransactionMode.Manual)]
    public class CreateToposolidFromRefPoints : IExternalCommand
    {
        private const string RefPointFamilyName = "RefPoint";

        // ====== CẤU HÌNH ======
        private const double GridSpacingMeters = 2.0;   // spacing lưới trong lòng địa hình
        private const double EdgeSpacingMeters = 1.0;   // spacing trên biên
        private const int IDW_K = 6;     // số lân cận trong IDW
        private const double IDW_Power = 2.0;   // lũy thừa IDW
        private const double SnapDistanceMeters = 0.05;  // ~5cm để “bắt” cao độ RefPoint
        private const int MaxPoints = 18000; // giới hạn hiệu năng

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // B1. Chọn Toposolid
                Reference topoRef = uidoc.Selection.PickObject(ObjectType.Element, new ToposolidFilter(), "Chọn Toposolid muốn chỉnh sửa");
                var originalTopo = doc.GetElement(topoRef) as Toposolid;
                if (originalTopo == null)
                {
                    TaskDialog.Show("Lỗi", "Không phải là Toposolid.");
                    return Result.Failed;
                }

                // B2. Gom RefPoint (Generic Model -> Family "RefPoint")
                var allRefFi = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Equals(RefPointFamilyName, StringComparison.InvariantCultureIgnoreCase) == true)
                    .ToList();

                if (allRefFi.Count == 0)
                {
                    TaskDialog.Show("Lỗi", "Không tìm thấy RefPoint nào.");
                    return Result.Failed;
                }

                var refPtsRaw = allRefFi
                    .Select(fi =>
                    {
                        var lp = fi.Location as LocationPoint;
                        if (lp == null) return null;

                        double z = lp.Point.Z;
                        var elevParam = fi.LookupParameter("Default Elevation");
                        if (elevParam != null && elevParam.StorageType == StorageType.Double)
                            z = elevParam.AsDouble();

                        return new XYZ(lp.Point.X, lp.Point.Y, z);
                    })
                    .Where(p => p != null)
                    .ToList();

                if (refPtsRaw.Count == 0)
                {
                    TaskDialog.Show("Lỗi", "RefPoint không hợp lệ.");
                    return Result.Failed;
                }

                // B3. Lấy biên & loops
                var topoBoundary = GetTopoBoundary(originalTopo);
                if (topoBoundary == null || topoBoundary.Count < 3)
                {
                    TaskDialog.Show("Lỗi", "Không lấy được biên dạng Toposolid.");
                    return Result.Failed;
                }

                var boundaryLoops = GetToposolidBoundary(originalTopo);
                if (boundaryLoops == null || boundaryLoops.Count == 0)
                {
                    TaskDialog.Show("Lỗi", "Không đọc được boundary của Toposolid.");
                    return Result.Failed;
                }

                // B4. Lấy type & level
                ElementId typeId = originalTopo.GetTypeId();
                ElementId levelId = originalTopo.LevelId;
                if (levelId == ElementId.InvalidElementId)
                {
                    var level = GetNearestLevel(doc, originalTopo);
                    if (level == null)
                    {
                        TaskDialog.Show("Lỗi", "Không tìm thấy Level.");
                        return Result.Failed;
                    }
                    levelId = level.Id;
                }

                // ====== TẠO TẬP ĐIỂM ======
                double gridSpacingFt = UnitUtils.ConvertToInternalUnits(GridSpacingMeters, UnitTypeId.Meters);
                double edgeSpacingFt = UnitUtils.ConvertToInternalUnits(EdgeSpacingMeters, UnitTypeId.Meters);
                double snapDistFt = UnitUtils.ConvertToInternalUnits(SnapDistanceMeters, UnitTypeId.Meters);

                // 1) Lưới đều trong polygon + jitter
                var gridPtsXY = SampleGridInsidePolygon(topoBoundary, gridSpacingFt, jitterRatio: 0.15);

                // 2) Densify biên
                var edgePtsXY = DensifyBoundary(topoBoundary, edgeSpacingFt);
                gridPtsXY.AddRange(edgePtsXY);

                // 3) Giữ các RefPoint bên trong polygon làm “neo” XY
                var insideRefXY = refPtsRaw.Where(p => IsPointInsidePolygonXY(p, topoBoundary))
                                           .Select(p => new XYZ(p.X, p.Y, 0.0))
                                           .ToList();
                gridPtsXY.AddRange(insideRefXY);

                // 4) Khử trùng lặp XY
                gridPtsXY = DeduplicateXY(gridPtsXY, tolerance: gridSpacingFt * 0.75);

                // 5) Nội suy Z bằng IDW từ tập RefPoint
                var elevatedPts = InterpolateZ_IDW(gridPtsXY, refPtsRaw, IDW_K, IDW_Power, snapDistFt);

                // 6) Giới hạn số lượng
                if (elevatedPts.Count > MaxPoints)
                    elevatedPts = Downsample(elevatedPts, MaxPoints);

                if (elevatedPts.Count == 0)
                {
                    TaskDialog.Show("Lỗi", "Không có điểm nào hợp lệ để tạo Toposolid.");
                    return Result.Failed;
                }

                // ====== TẠO LẠI TOPOSOLID ======
                using (var t = new Transaction(doc, "Cập nhật Toposolid từ RefPoint (mượt hoá)"))
                {
                    t.Start();

                    Toposolid newTopo = Toposolid.Create(doc, boundaryLoops, elevatedPts, typeId, levelId);

                    // Xoá Toposolid cũ
                    doc.Delete(originalTopo.Id);

                    t.Commit();
                }

                TaskDialog.Show("Xong", $"Tạo Toposolid mới với {elevatedPts.Count} điểm (grid + edge + ref).");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ================== SELECTION FILTER ==================
        private class ToposolidFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) =>
                elem?.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_Toposolid;
            public bool AllowReference(Reference reference, XYZ position) => true;
        }

        // ================== BOUNDARY HELPERS ==================
        private IList<XYZ> GetTopoBoundary(Toposolid topo)
        {
            var boundary = new List<XYZ>();
            var opt = new Options
            {
                ComputeReferences = false,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = false
            };

            var geomElem = topo.get_Geometry(opt);
            if (geomElem == null) return boundary;

            foreach (GeometryObject obj in geomElem)
            {
                if (obj is Solid solid && solid.Faces.Size > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace pf && pf.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ.Negate()))
                        {
                            var loops = face.GetEdgesAsCurveLoops();
                            if (loops.Count > 0)
                            {
                                var outer = loops.OrderByDescending(l => l.GetExactLength()).First();
                                foreach (Curve c in outer)
                                    boundary.Add(c.GetEndPoint(0));
                                return boundary;
                            }
                        }
                    }
                }
            }
            return boundary;
        }

        private IList<CurveLoop> GetToposolidBoundary(Toposolid topo)
        {
            var opt = new Options
            {
                ComputeReferences = false,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = false
            };

            var geomElem = topo.get_Geometry(opt);
            if (geomElem == null) return null;

            foreach (GeometryObject obj in geomElem)
            {
                if (obj is Solid solid && solid.Faces.Size > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace pf && pf.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ.Negate()))
                            return face.GetEdgesAsCurveLoops();
                    }
                }
            }
            return null;
        }

        private Level GetNearestLevel(Document doc, Toposolid topo)
        {
            var bbox = topo.get_BoundingBox(null);
            if (bbox == null) return null;

            double z = bbox.Min.Z;
            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => Math.Abs(l.Elevation - z))
                .ToList();
            return levels.FirstOrDefault();
        }

        // ================== SPATIAL HELPERS ==================
        private bool IsPointInsidePolygonXY(XYZ p, IList<XYZ> polygon)
        {
            int n = polygon.Count;
            bool inside = false;
            for (int i = 0, j = n - 1; i < n; j = i++)
            {
                double xi = polygon[i].X, yi = polygon[i].Y;
                double xj = polygon[j].X, yj = polygon[j].Y;

                bool intersect = ((yi > p.Y) != (yj > p.Y)) &&
                                 (p.X < (xj - xi) * (p.Y - yi) / ((yj - yi) == 0 ? 1e-12 : (yj - yi)) + xi);
                if (intersect) inside = !inside;
            }
            return inside;
        }

        private List<XYZ> DensifyBoundary(IList<XYZ> boundary, double spacing)
        {
            var pts = new List<XYZ>();
            int n = boundary.Count;
            for (int i = 0; i < n; i++)
            {
                XYZ a = boundary[i];
                XYZ b = boundary[(i + 1) % n];
                double len = a.DistanceTo(b);
                int segments = Math.Max(1, (int)Math.Round(len / spacing));

                for (int s = 0; s <= segments; s++)
                {
                    double t = (double)s / segments;
                    XYZ p = a + (b - a) * t;
                    pts.Add(new XYZ(p.X, p.Y, 0.0)); // Z sẽ nội suy
                }
            }
            return pts;
        }

        private List<XYZ> SampleGridInsidePolygon(IList<XYZ> poly, double spacing, double jitterRatio = 0.1)
        {
            var pts = new List<XYZ>();
            double minX = poly.Min(p => p.X), maxX = poly.Max(p => p.X);
            double minY = poly.Min(p => p.Y), maxY = poly.Max(p => p.Y);

            Random rnd = new Random(12345);
            double jitter = spacing * jitterRatio;

            int rowIndex = 0;
            for (double y = minY; y <= maxY; y += spacing, rowIndex++)
            {
                double rowOffset = (rowIndex % 2 == 0) ? 0 : spacing * 0.5; // pseudo-hex
                for (double x = minX + rowOffset; x <= maxX; x += spacing)
                {
                    double jx = (rnd.NextDouble() - 0.5) * 2.0 * jitter;
                    double jy = (rnd.NextDouble() - 0.5) * 2.0 * jitter;
                    var p = new XYZ(x + jx, y + jy, 0.0);
                    if (IsPointInsidePolygonXY(p, poly))
                        pts.Add(p);
                }
            }
            return pts;
        }

        private List<XYZ> DeduplicateXY(IEnumerable<XYZ> pts, double tolerance)
        {
            var snap = new Dictionary<(int, int), XYZ>();
            double inv = 1.0 / Math.Max(tolerance, 1e-9);

            foreach (var p in pts)
            {
                int ix = (int)Math.Round(p.X * inv);
                int iy = (int)Math.Round(p.Y * inv);
                var key = (ix, iy);
                if (!snap.ContainsKey(key))
                    snap[key] = p;
            }
            return snap.Values.ToList();
        }

        private List<XYZ> Downsample(List<XYZ> pts, int maxCount)
        {
            Random rnd = new Random(6789);
            for (int i = pts.Count - 1; i > 0; i--)
            {
                int j = rnd.Next(i + 1);
                (pts[i], pts[j]) = (pts[j], pts[i]);
            }
            return pts.Take(maxCount).ToList();
        }

        // ================== IDW ==================
        private List<XYZ> InterpolateZ_IDW(List<XYZ> targetsXY, List<XYZ> refPts, int k, double power, double snapDist)
        {
            var outPts = new List<XYZ>(targetsXY.Count);

            foreach (var q in targetsXY)
            {
                var neighbors = refPts
                    .Select(r => new { R = r, D = DistXY(q, r) })
                    .OrderBy(x => x.D)
                    .Take(Math.Max(1, k))
                    .ToList();

                if (neighbors.Count == 0)
                {
                    outPts.Add(new XYZ(q.X, q.Y, 0));
                    continue;
                }

                if (neighbors[0].D < snapDist)
                {
                    outPts.Add(new XYZ(q.X, q.Y, neighbors[0].R.Z));
                    continue;
                }

                double wsum = 0.0;
                double zsum = 0.0;

                foreach (var nb in neighbors)
                {
                    double d = Math.Max(nb.D, 1e-9);
                    double w = 1.0 / Math.Pow(d, power);
                    wsum += w;
                    zsum += w * nb.R.Z;
                }

                double zq = (wsum > 0) ? (zsum / wsum) : neighbors[0].R.Z;
                outPts.Add(new XYZ(q.X, q.Y, zq));
            }

            return outPts;
        }

        private double DistXY(XYZ a, XYZ b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }
    }
}
