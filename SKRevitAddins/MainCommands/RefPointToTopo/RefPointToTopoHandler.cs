using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.RefPointToTopo
{
    /// <summary>
    /// Chạy trên Revit API thread: build điểm (grid + edge + adaptive), nội suy IDW, tạo Toposolid mới và xóa cái cũ.
    /// ViewModel set các thuộc tính rồi gọi ExternalEvent.Raise().
    /// </summary>
    public class RefPointToTopoHandler : IExternalEventHandler
    {
        // ====== INPUT từ ViewModel ======
        public ElementId TargetToposolidId { get; set; }
        public string RefPointFamilyName { get; set; } = "RefPoint";

        // Config (đơn vị FEET)
        public double GridSpacingFt { get; set; } = 2.0;   // convert từ mét trước khi gán
        public double EdgeSpacingFt { get; set; } = 1.0;
        public int IDW_K { get; set; } = 6;
        public double IDW_Power { get; set; } = 2.0;
        public double SnapDistanceFt { get; set; } = 0.05;
        public int MaxPoints { get; set; } = 18000;

        // Adaptive
        public bool UseAdaptiveSampling { get; set; } = true;
        public double GradThreshold { get; set; } = 0.30;
        public double RefineFactor { get; set; } = 0.5;
        public int MaxRefinePoints { get; set; } = 6000;

        // Hiệu năng
        public bool UseParallelIDW { get; set; } = true;

        // UI hooks
        public bool IsCancelled { get; set; } = false;
        public Action<bool> BusySetter { get; set; }
        public Action<int, int> ProgressReporter { get; set; }  // (current, total)

        public void Execute(UIApplication app)
        {
            try
            {
                BusySetter?.Invoke(true);
                ProgressReporter?.Invoke(0, 1);

                var uidoc = app.ActiveUIDocument;
                var doc = uidoc.Document;

                var originalTopo = doc.GetElement(TargetToposolidId) as Toposolid;
                if (originalTopo == null)
                {
                    TaskDialog.Show("RefPoint → Toposolid", "Toposolid không hợp lệ.");
                    return;
                }

                // ====== B2. Gom RefPoint ======
                var allRefFi = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Equals(RefPointFamilyName, StringComparison.InvariantCultureIgnoreCase) == true)
                    .ToList();

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
                    TaskDialog.Show("RefPoint → Toposolid", "Không tìm thấy RefPoint hợp lệ.");
                    return;
                }

                // Lọc outlier theo Z-score
                refPtsRaw = RemoveOutliers(refPtsRaw, 3.0);

                // ====== B3. Boundary ======
                var boundaryLoops = GetToposolidBoundary(originalTopo);
                if (boundaryLoops == null || boundaryLoops.Count == 0)
                {
                    TaskDialog.Show("RefPoint → Toposolid", "Không đọc được boundary loops.");
                    return;
                }

                // polygon outer (dùng để test Inside)
                var topoBoundary = GetTopoBoundary(originalTopo);
                if (topoBoundary == null || topoBoundary.Count < 3)
                {
                    TaskDialog.Show("RefPoint → Toposolid", "Không lấy được biên Toposolid.");
                    return;
                }

                // ====== B4. Type & Level ======
                ElementId typeId = originalTopo.GetTypeId();
                ElementId levelId = originalTopo.LevelId;
                if (levelId == ElementId.InvalidElementId)
                {
                    var level = GetNearestLevel(doc, originalTopo);
                    if (level == null)
                    {
                        TaskDialog.Show("RefPoint → Toposolid", "Không tìm thấy Level.");
                        return;
                    }
                    levelId = level.Id;
                }

                // ====== Sampling XY ======
                var gridPtsXY = SampleGridInsidePolygon(topoBoundary, GridSpacingFt, 0.15);
                var edgePtsXY = DensifyBoundary(topoBoundary, EdgeSpacingFt);
                gridPtsXY.AddRange(edgePtsXY);

                // thêm các RefPoint nằm trong polygon làm neo XY
                var insideRefXY = refPtsRaw.Where(p => IsPointInsidePolygonXY(p, topoBoundary))
                                           .Select(p => new XYZ(p.X, p.Y, 0.0))
                                           .ToList();
                gridPtsXY.AddRange(insideRefXY);

                // Khử trùng lặp XY
                gridPtsXY = DeduplicateXY(gridPtsXY, Math.Max(GridSpacingFt * 0.75, 1e-6));

                // ====== IDW lần 1 ======
                var elevatedPts = InterpolateZ_IDW(gridPtsXY, refPtsRaw, IDW_K, IDW_Power, SnapDistanceFt, UseParallelIDW);

                // ====== Adaptive refine ======
                if (UseAdaptiveSampling)
                {
                    var refineXY = AdaptiveRefine(elevatedPts, GridSpacingFt, topoBoundary, GradThreshold, RefineFactor, MaxRefinePoints);
                    if (refineXY.Count > 0)
                    {
                        var refineElev = InterpolateZ_IDW(refineXY, refPtsRaw, IDW_K, IDW_Power, SnapDistanceFt, UseParallelIDW);
                        elevatedPts.AddRange(refineElev);
                        elevatedPts = DeduplicateXY(elevatedPts, Math.Min(GridSpacingFt * 0.5, EdgeSpacingFt * 0.75));
                    }
                }

                if (elevatedPts.Count > MaxPoints)
                    elevatedPts = Downsample(elevatedPts, MaxPoints);

                if (elevatedPts.Count == 0)
                {
                    TaskDialog.Show("RefPoint → Toposolid", "Không có điểm hợp lệ để tạo Toposolid.");
                    return;
                }

                // Clamp Z theo dải gốc
                double zMin = refPtsRaw.Min(p => p.Z);
                double zMax = refPtsRaw.Max(p => p.Z);
                for (int i = 0; i < elevatedPts.Count; i++)
                {
                    var p = elevatedPts[i];
                    double z = Math.Max(zMin, Math.Min(zMax, p.Z));
                    elevatedPts[i] = new XYZ(p.X, p.Y, z);
                }

                // ====== Tạo lại Toposolid ======
                using (var t = new Transaction(doc, "Cập nhật Toposolid từ RefPoint (MVVM)"))
                {
                    t.Start();

                    Toposolid.Create(doc, boundaryLoops, elevatedPts, typeId, levelId);
                    doc.Delete(originalTopo.Id);

                    t.Commit();
                }

                TaskDialog.Show("Xong", $"Tạo Toposolid mới với {elevatedPts.Count} điểm (grid + edge + adaptive).");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RefPoint → Toposolid (Error)", ex.Message);
            }
            finally
            {
                BusySetter?.Invoke(false);
                IsCancelled = false;
                ProgressReporter?.Invoke(1, 1);
            }
        }

        public string GetName() => "RefPointToTopo Handler";

        // ================== GEOMETRY / BOUNDARY ==================
        private IList<XYZ> GetTopoBoundary(Toposolid topo)
        {
            var boundary = new List<XYZ>();
            var opt = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Fine, IncludeNonVisibleObjects = false };
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
                            var loops = pf.GetEdgesAsCurveLoops();
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
            var opt = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Fine, IncludeNonVisibleObjects = false };
            var geomElem = topo.get_Geometry(opt);
            if (geomElem == null) return null;

            foreach (GeometryObject obj in geomElem)
            {
                if (obj is Solid solid && solid.Faces.Size > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace pf && pf.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ.Negate()))
                            return face.GetEdgesAsCurveLoops(); // gồm cả outer + inner
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
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => Math.Abs(l.Elevation - z))
                .FirstOrDefault();
        }

        // ================== SAMPLING / IDW ==================
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
                    pts.Add(new XYZ(p.X, p.Y, 0.0));
                }
            }
            return pts;
        }

        private List<XYZ> SampleGridInsidePolygon(IList<XYZ> poly, double spacing, double jitterRatio = 0.1)
        {
            var pts = new List<XYZ>();
            double minX = poly.Min(p => p.X), maxX = poly.Max(p => p.X);
            double minY = poly.Min(p => p.Y), maxY = poly.Max(p => p.Y);

            var rnd = new Random(12345);
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
            var rnd = new Random(6789);
            for (int i = pts.Count - 1; i > 0; i--)
            {
                int j = rnd.Next(i + 1);
                (pts[i], pts[j]) = (pts[j], pts[i]);
            }
            return pts.Take(maxCount).ToList();
        }

        private List<XYZ> RemoveOutliers(List<XYZ> pts, double zThresh = 3.0)
        {
            if (pts.Count < 5) return pts;
            double mean = pts.Average(p => p.Z);
            double var = pts.Average(p => (p.Z - mean) * (p.Z - mean));
            double sd = Math.Sqrt(Math.Max(var, 1e-12));
            return pts.Where(p => Math.Abs(p.Z - mean) <= zThresh * sd).ToList();
        }

        private List<XYZ> AdaptiveRefine(
            List<XYZ> elevated, double baseSpacing, IList<XYZ> polygon,
            double gradThreshold, double refineFactor, int maxRefine)
        {
            var refine = new List<XYZ>();
            if (elevated.Count == 0) return refine;

            // Gom theo ô lưới thưa
            var groups = elevated.GroupBy(p => (Math.Floor(p.X / baseSpacing), Math.Floor(p.Y / baseSpacing)));
            double rSpacing = Math.Max(baseSpacing * refineFactor, baseSpacing * 0.33); // tránh quá nhỏ

            foreach (var g in groups)
            {
                var cell = g.ToList();
                if (cell.Count < 3) continue;

                double zmin = cell.Min(p => p.Z);
                double zmax = cell.Max(p => p.Z);
                double grad = (zmax - zmin) / Math.Max(baseSpacing, 1e-9);

                if (grad > gradThreshold)
                {
                    double cx = cell.Average(p => p.X);
                    double cy = cell.Average(p => p.Y);
                    double d = rSpacing * 0.5;

                    var cands = new[]
                    {
                        new XYZ(cx + d, cy,     0),
                        new XYZ(cx - d, cy,     0),
                        new XYZ(cx,     cy + d, 0),
                        new XYZ(cx,     cy - d, 0)
                    };
                    foreach (var q in cands)
                        if (IsPointInsidePolygonXY(q, polygon)) refine.Add(q);
                }
            }

            if (refine.Count == 0) return refine;
            refine = DeduplicateXY(refine, rSpacing * 0.6);
            if (refine.Count > maxRefine) refine = Downsample(refine, maxRefine);
            return refine;
        }

        private List<XYZ> InterpolateZ_IDW(List<XYZ> targetsXY, List<XYZ> refPts, int k, double power, double snapDist, bool parallel)
        {
            if (!parallel) return InterpolateZ_IDW_Serial(targetsXY, refPts, k, power, snapDist);
            return InterpolateZ_IDW_Parallel(targetsXY, refPts, k, power, snapDist);
        }

        private List<XYZ> InterpolateZ_IDW_Serial(List<XYZ> targetsXY, List<XYZ> refPts, int k, double power, double snapDist)
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

        private List<XYZ> InterpolateZ_IDW_Parallel(List<XYZ> targetsXY, List<XYZ> refPts, int k, double power, double snapDist)
        {
            var outArr = new XYZ[targetsXY.Count];

            Parallel.For(0, targetsXY.Count, i =>
            {
                var q = targetsXY[i];
                var neighbors = refPts
                    .Select(r => new { R = r, D = DistXY(q, r) })
                    .OrderBy(x => x.D)
                    .Take(Math.Max(1, k))
                    .ToList();

                if (neighbors.Count == 0)
                {
                    outArr[i] = new XYZ(q.X, q.Y, 0);
                    return;
                }

                if (neighbors[0].D < snapDist)
                {
                    outArr[i] = new XYZ(q.X, q.Y, neighbors[0].R.Z);
                    return;
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
                outArr[i] = new XYZ(q.X, q.Y, zq);
            });

            return outArr.ToList();
        }

        private double DistXY(XYZ a, XYZ b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }
    }
}
