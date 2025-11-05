using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Threading.Tasks;
using static System.Math;

namespace SKRevitAddins.PointCloudAddins.RefPointToTopo
{
    public class RefPointToTopoHandler : IExternalEventHandler
    {
        public ElementId TargetToposolidId { get; set; }
        public string RefPointFamilyName { get; set; } = "RefPoint";
        public double GridSpacingFt { get; set; } = 2.0;
        public double EdgeSpacingFt { get; set; } = 1.0;
        public int IDW_K { get; set; } = 6;
        public double IDW_Power { get; set; } = 2.0;
        public double SnapDistanceFt { get; set; } = 0.05;
        public int MaxPoints { get; set; } = 18000;
        public bool UseAdaptiveSampling { get; set; } = true;
        public double GradThreshold { get; set; } = 0.30;
        public double RefineFactor { get; set; } = 0.5;
        public int MaxRefinePoints { get; set; } = 6000;
        public bool UseParallelIDW { get; set; } = true;
        public bool IsCancelled { get; set; } = false;
        public Action<bool> BusySetter { get; set; }
        public Action<int, int> ProgressReporter { get; set; }

        // Thêm thuộc tính CreateTopoFromRefPointOnly
        public bool CreateTopoFromRefPointOnly { get; set; } = false;

        public void Execute(UIApplication app)
        {
            int step = 0, total = 11;
            Action tick = () => ProgressReporter?.Invoke(step < total ? ++step : total, total);

            try
            {
                BusySetter?.Invoke(true);
                ProgressReporter?.Invoke(0, total);

                var doc = app.ActiveUIDocument.Document;
                var originalTopo = doc.GetElement(TargetToposolidId) as Toposolid;
                if (originalTopo == null)
                {
                    TaskDialog.Show("RefPoint → Toposolid", "Toposolid không hợp lệ.");
                    return;
                }

                // ===== 1) Thu thập RefPoint =====
                if (IsCancelled) return;
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
                        foreach (var name in new[] { "Default Elevation", "Elevation" })
                        {
                            var p = fi.LookupParameter(name);
                            if (p != null && p.StorageType == StorageType.Double) { z = p.AsDouble(); break; }
                        }

                        return new XYZ(lp.Point.X, lp.Point.Y, z);
                    })
                    .Where(p => p != null)
                    .ToList();

                if (refPtsRaw.Count == 0)
                {
                    TaskDialog.Show("RefPoint → Toposolid", "Không tìm thấy RefPoint hợp lệ.");
                    return;
                }
                tick();

                // Lọc outlier theo Z-score
                refPtsRaw = RemoveOutliers(refPtsRaw, 3.0);

                // ===== 2) Boundary / Loops =====
                if (IsCancelled) return;
                var boundaryLoops = GetToposolidBoundary(originalTopo);
                if (boundaryLoops == null || boundaryLoops.Count == 0)
                {
                    TaskDialog.Show("RefPoint → Toposolid", "Không đọc được boundary loops.");
                    return;
                }
                var topoBoundary = GetTopoBoundary(originalTopo);
                if (topoBoundary == null || topoBoundary.Count < 3)
                {
                    TaskDialog.Show("RefPoint → Toposolid", "Không lấy được biên Toposolid.");
                    return;
                }
                tick();

                // ===== 3) Type & Level =====
                if (IsCancelled) return;
                ElementId typeId = originalTopo.GetTypeId();
                ElementId levelId = originalTopo.LevelId;
                if (levelId == ElementId.InvalidElementId)
                {
                    var nearestLevel = GetNearestLevel(doc, originalTopo);
                    if (nearestLevel == null)
                    {
                        TaskDialog.Show("RefPoint → Toposolid", "Không tìm thấy Level.");
                        return;
                    }
                    levelId = nearestLevel.Id;
                }
                tick();

                // ===== 4) Sampling XY =====
                var gridPtsXY = new List<XYZ>();

                // Nếu chỉ tạo topo từ RefPoint
                if (CreateTopoFromRefPointOnly)
                {
                    gridPtsXY.AddRange(refPtsRaw.Select(p => new XYZ(p.X, p.Y, 0.0)));
                }
                else
                {
                    var edgePtsXY = DensifyBoundary(topoBoundary, EdgeSpacingFt, () => IsCancelled);
                    gridPtsXY.AddRange(edgePtsXY);

                    var insideRefPts = refPtsRaw.Where(p => IsPointInsidePolygonXY(p, topoBoundary)).ToList();
                    var insideRefXY = insideRefPts.Select(p => new XYZ(p.X, p.Y, 0.0)).ToList();
                    gridPtsXY.AddRange(insideRefXY);

                    gridPtsXY = DeduplicateXY(gridPtsXY, Max(GridSpacingFt * 0.75, 1e-6));
                }
                tick();

                // ===== 5) IDW lần 1 =====
                if (IsCancelled) return;
                var elevatedPts = InterpolateZ_IDW(gridPtsXY, refPtsRaw, IDW_K, IDW_Power, SnapDistanceFt, UseParallelIDW, () => IsCancelled);
                tick();

                // ===== 6) Adaptive refine =====
                if (UseAdaptiveSampling && !IsCancelled)
                {
                    var refineXY = AdaptiveRefine(elevatedPts, GridSpacingFt, topoBoundary, GradThreshold, RefineFactor, MaxRefinePoints, () => IsCancelled);
                    if (refineXY.Count > 0)
                    {
                        var refineElev = InterpolateZ_IDW(refineXY, refPtsRaw, IDW_K, IDW_Power, SnapDistanceFt, UseParallelIDW, () => IsCancelled);
                        elevatedPts.AddRange(refineElev);
                        elevatedPts = DeduplicateXY(elevatedPts, Min(GridSpacingFt * 0.5, EdgeSpacingFt * 0.75));
                    }
                }
                tick();

                // ===== 7) Merge ưu tiên RefPoint =====
                if (IsCancelled) return;
                double preferTol = Min(GridSpacingFt * 0.5, EdgeSpacingFt * 0.75);
                elevatedPts = MergePreferRefPoints(elevatedPts, refPtsRaw, preferTol);
                tick();

                // ===== 8) Giới hạn số điểm =====
                elevatedPts = EnforceMaxPointsWithRefPriority(elevatedPts, refPtsRaw, MaxPoints, preferTol);
                if (elevatedPts.Count == 0)
                {
                    TaskDialog.Show("RefPoint → Toposolid", "Không có điểm hợp lệ để tạo Toposolid.");
                    return;
                }
                tick();

                // ===== 9) Clamp Z CHỈ cho điểm nội suy (KHÔNG clamp RefPoint) =====
                if (IsCancelled) return;
                double zMin = refPtsRaw.Min(p => p.Z);
                double zMax = refPtsRaw.Max(p => p.Z);

                double inv = 1.0 / Max(preferTol, 1e-9);
                var refKeys = new HashSet<(int, int)>(
                    refPtsRaw.Select(r => ((int)Round(r.X * inv), (int)Round(r.Y * inv)))
                );

                for (int i = 0; i < elevatedPts.Count; i++)
                {
                    var p = elevatedPts[i];
                    var key = ((int)Round(p.X * inv), (int)Round(p.Y * inv));

                    if (!refKeys.Contains(key))
                    {
                        // Điểm nội suy → clamp theo min/max RefPoint
                        double z = Max(zMin, Min(zMax, p.Z));
                        elevatedPts[i] = new XYZ(p.X, p.Y, z);
                    }
                    // else: RefPoint giữ nguyên
                    if (IsCancelled) return;
                }
                tick();

                // ===== 10) Force include RefPoint =====
                elevatedPts = ForceIncludeRefPoints(elevatedPts, refPtsRaw, preferTol);
                tick();

                // ===== 10.5) Sửa "too thin": chỉ offset điểm NỘI SUY =====
                if (IsCancelled) return;
                var levelElem = (Level)doc.GetElement(levelId);
                double levelElev = levelElem?.Elevation ?? 0.0;

                double thickness = 0.0;
                try
                {
                    var tType = doc.GetElement(typeId) as ToposolidType;
                    var cs = tType?.GetCompoundStructure();
                    if (cs != null)
                    {
                        thickness = cs.GetWidth();
                    }
                }
                catch { thickness = 0.0; }

                double eps = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters); // 1 mm
                double minAllowed = levelElev + thickness + eps;

                double minInterpZ = double.PositiveInfinity;
                double minRefZ = double.PositiveInfinity;

                foreach (var p in elevatedPts)
                {
                    var key = ((int)Round(p.X * inv), (int)Round(p.Y * inv));
                    if (refKeys.Contains(key))
                    {
                        if (p.Z < minRefZ) minRefZ = p.Z;
                    }
                    else
                    {
                        if (p.Z < minInterpZ) minInterpZ = p.Z;
                    }
                }

                // 1) Nếu nội suy vi phạm, chỉ nâng nội suy
                if (minInterpZ < minAllowed && minRefZ >= minAllowed)
                {
                    double dz = minAllowed - minInterpZ;
                    for (int i = 0; i < elevatedPts.Count; i++)
                    {
                        var p = elevatedPts[i];
                        var key = ((int)Round(p.X * inv), (int)Round(p.Y * inv));
                        if (!refKeys.Contains(key))
                            elevatedPts[i] = new XYZ(p.X, p.Y, p.Z + dz);
                    }
                }
                // 2) Nếu chính RefPoint vi phạm, không tự nâng (để giữ Z RefPoint)
                else if (minRefZ < minAllowed)
                {
                    TaskDialog.Show("RefPoint → Toposolid",
                        "Một hoặc nhiều RefPoint có Z thấp hơn bề dày tối thiểu của loại Toposolid.\n" +
                        "Để giữ Z RefPoint, hãy chọn loại Toposolid mỏng hơn hoặc điều chỉnh Level/offset.");
                }
                tick();

                // ===== 11) Transaction an toàn =====
                using (var t = new Transaction(doc, "Cập nhật Toposolid từ RefPoint (MVVM)"))
                {
                    t.Start();
                    Toposolid newTopo = null;
                    try
                    {
                        newTopo = Toposolid.Create(doc, boundaryLoops, elevatedPts, typeId, levelId);
                        if (newTopo == null) throw new InvalidOperationException("Không tạo được Toposolid mới.");
                        doc.Delete(originalTopo.Id);
                        t.Commit();
                    }
                    catch
                    {
                        if (t.HasStarted()) t.RollBack();
                        throw;
                    }
                }
                tick();

                TaskDialog.Show("Xong", $"Tạo Toposolid mới với {elevatedPts.Count} điểm (mọi RefPoint đều có subelement, Z RefPoint giữ nguyên).");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RefPoint → Toposolid (Lỗi)", ex.Message);
            }
            finally
            {
                BusySetter?.Invoke(false);
                IsCancelled = false;
                ProgressReporter?.Invoke(total, total);
            }
        }

        public string GetName() => "RefPointToTopo Handler";

        // Các helper methods

        private List<XYZ> RemoveOutliers(List<XYZ> pts, double zThresh = 3.0)
        {
            if (pts.Count < 5) return pts;
            double mean = pts.Average(p => p.Z);
            double var = pts.Average(p => (p.Z - mean) * (p.Z - mean));
            double sd = Math.Sqrt(Math.Max(var, 1e-12));
            return pts.Where(p => Math.Abs(p.Z - mean) <= zThresh * sd).ToList();
        }

        private List<XYZ> DensifyBoundary(IList<XYZ> boundary, double spacing, Func<bool> isCancelled)
        {
            var pts = new List<XYZ>();
            int n = boundary.Count;
            for (int i = 0; i < n; i++)
            {
                var a = boundary[i];
                var b = boundary[(i + 1) % n];
                double len = a.DistanceTo(b);
                int seg = Math.Max(1, (int)Math.Round(len / spacing));

                for (int s = 0; s <= seg; s++)
                {
                    double t = (double)s / seg;
                    var p = a + (b - a) * t;
                    pts.Add(new XYZ(p.X, p.Y, 0.0));

                    if (isCancelled()) return pts;
                }
            }
            return pts;
        }

        // Phương thức DeduplicateXY mà không có tham số tolerance
        private List<XYZ> DeduplicateXYWithoutTolerance(IEnumerable<XYZ> pts)
        {
            var map = new Dictionary<(int, int), XYZ>();
            double inv = 1.0 / Math.Max(1e-6, 1e-9); // Giả sử đây là giá trị mặc định nếu không có tolerance

            foreach (var p in pts)
            {
                int ix = (int)Math.Round(p.X * inv);
                int iy = (int)Math.Round(p.Y * inv);
                var key = (ix, iy);
                if (!map.ContainsKey(key)) map[key] = p;
            }
            return map.Values.ToList();
        }

        // Phương thức DeduplicateXY với tolerance
        private List<XYZ> DeduplicateXYWithTolerance(IEnumerable<XYZ> pts, double tol)
        {
            var map = new Dictionary<(int, int), XYZ>();
            double inv = 1.0 / Math.Max(tol, 1e-9);

            foreach (var p in pts)
            {
                int ix = (int)Math.Round(p.X * inv);
                int iy = (int)Math.Round(p.Y * inv);
                var key = (ix, iy);
                if (!map.ContainsKey(key)) map[key] = p;
            }
            return map.Values.ToList();
        }


        private List<XYZ> InterpolateZ_IDW(List<XYZ> targetsXY, List<XYZ> refPts, int k, double power, double snap, bool parallel, Func<bool> isCancelled)
        {
            return parallel ? InterpolateZ_IDW_Parallel(targetsXY, refPts, k, power, snap, isCancelled)
                            : InterpolateZ_IDW_Serial(targetsXY, refPts, k, power, snap, isCancelled);
        }

        private List<XYZ> InterpolateZ_IDW_Serial(List<XYZ> targetsXY, List<XYZ> refPts, int k, double power, double snap, Func<bool> isCancelled)
        {
            var outPts = new List<XYZ>(targetsXY.Count);
            foreach (var q in targetsXY)
            {
                var nbs = refPts.Select(r => new { R = r, D = DistXY(q, r) })
                                .OrderBy(x => x.D).Take(Math.Max(1, k)).ToList();

                if (nbs.Count == 0) { outPts.Add(new XYZ(q.X, q.Y, 0)); continue; }
                if (nbs[0].D < snap) { outPts.Add(new XYZ(q.X, q.Y, nbs[0].R.Z)); continue; }

                double wsum = 0, zsum = 0;
                foreach (var nb in nbs)
                {
                    double d = Math.Max(nb.D, 1e-9);
                    double w = 1.0 / Math.Pow(d, power);
                    wsum += w; zsum += w * nb.R.Z;
                }
                double zq = (wsum > 0) ? (zsum / wsum) : nbs[0].R.Z;
                outPts.Add(new XYZ(q.X, q.Y, zq));

                if (isCancelled()) return outPts;
            }
            return outPts;
        }

        private List<XYZ> InterpolateZ_IDW_Parallel(List<XYZ> targetsXY, List<XYZ> refPts, int k, double power, double snap, Func<bool> isCancelled)
        {
            var outArr = new XYZ[targetsXY.Count];
            Parallel.For(0, targetsXY.Count, i =>
            {
                var q = targetsXY[i];
                var nbs = refPts.Select(r => new { R = r, D = DistXY(q, r) })
                                .OrderBy(x => x.D).Take(Math.Max(1, k)).ToList();

                if (nbs.Count == 0) { outArr[i] = new XYZ(q.X, q.Y, 0); return; }
                if (nbs[0].D < snap) { outArr[i] = new XYZ(q.X, q.Y, nbs[0].R.Z); return; }

                double wsum = 0, zsum = 0;
                foreach (var nb in nbs)
                {
                    double d = Math.Max(nb.D, 1e-9);
                    double w = 1.0 / Math.Pow(d, power);
                    wsum += w; zsum += w * nb.R.Z;
                }
                double zq = (wsum > 0) ? (zsum / wsum) : nbs[0].R.Z;
                outArr[i] = new XYZ(q.X, q.Y, zq);

                if (isCancelled()) return;
            });
            return outArr.ToList();
        }
        // 1. Lấy boundary của Toposolid
        private IList<XYZ> GetTopoBoundary(Toposolid topo)
        {
            var boundary = new List<XYZ>();
            var opt = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Fine, IncludeNonVisibleObjects = false };
            var ge = topo.get_Geometry(opt);
            if (ge == null) return boundary;

            foreach (GeometryObject obj in ge)
            {
                if (obj is Solid s && s.Faces.Size > 0)
                {
                    foreach (Face f in s.Faces)
                    {
                        if (f is PlanarFace pf && pf.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ.Negate()))
                        {
                            var loops = pf.GetEdgesAsCurveLoops();
                            if (loops.Count > 0)
                            {
                                var outer = loops.OrderByDescending(l => l.GetExactLength()).First();
                                foreach (Curve c in outer) boundary.Add(c.GetEndPoint(0));
                                return boundary;
                            }
                        }
                    }
                }
            }
            return boundary;
        }

        // 2. Lấy boundary Toposolid dạng CurveLoop
        private IList<CurveLoop> GetToposolidBoundary(Toposolid topo)
        {
            var opt = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Fine, IncludeNonVisibleObjects = false };
            var ge = topo.get_Geometry(opt);
            if (ge == null) return null;

            foreach (GeometryObject obj in ge)
            {
                if (obj is Solid s && s.Faces.Size > 0)
                {
                    foreach (Face f in s.Faces)
                        if (f is PlanarFace pf && pf.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ.Negate()))
                            return pf.GetEdgesAsCurveLoops();
                }
            }
            return null;
        }

        // 3. Lấy Level gần nhất với Toposolid
        private Level GetNearestLevel(Document doc, Toposolid topo)
        {
            var bb = topo.get_BoundingBox(null); if (bb == null) return null;
            double z = bb.Min.Z;
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level)).Cast<Level>()
                .OrderBy(l => Math.Abs(l.Elevation - z)).FirstOrDefault();
        }

        // 4. Kiểm tra nếu điểm nằm trong polygon XY
        private bool IsPointInsidePolygonXY(XYZ p, IList<XYZ> poly)
        {
            int n = poly.Count; bool inside = false;
            for (int i = 0, j = n - 1; i < n; j = i++)
            {
                double xi = poly[i].X, yi = poly[i].Y;
                double xj = poly[j].X, yj = poly[j].Y;

                bool intersect = ((yi > p.Y) != (yj > p.Y)) &&
                                 (p.X < (xj - xi) * (p.Y - yi) / ((yj - yi) == 0 ? 1e-12 : (yj - yi)) + xi);
                if (intersect) inside = !inside;
            }
            return inside;
        }

        // 5. Phương pháp AdaptiveRefine
        private List<XYZ> AdaptiveRefine(List<XYZ> elevated, double baseSpacing, IList<XYZ> polygon,
                                         double gradThreshold, double refineFactor, int maxRefine, Func<bool> isCancelled)
        {
            var refine = new List<XYZ>();
            if (elevated.Count == 0) return refine;

            var groups = elevated.GroupBy(p => (Math.Floor(p.X / baseSpacing), Math.Floor(p.Y / baseSpacing)));
            double rSpacing = Math.Max(baseSpacing * refineFactor, baseSpacing * 0.33);

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

                    var cands = new[] { new XYZ(cx + d, cy, 0), new XYZ(cx - d, cy, 0), new XYZ(cx, cy + d, 0), new XYZ(cx, cy - d, 0) };
                    foreach (var q in cands)
                        if (IsPointInsidePolygonXY(q, polygon)) refine.Add(q);
                }

                if (isCancelled()) return refine;
            }

            if (refine.Count == 0) return refine;
            refine = DeduplicateXY(refine, rSpacing * 0.6);
            if (refine.Count > maxRefine) refine = Downsample(refine, maxRefine);
            return refine;
        }

        // 6. Merge ưu tiên các RefPoint
        private List<XYZ> MergePreferRefPoints(List<XYZ> computed, List<XYZ> refPts, double tol)
        {
            var map = new Dictionary<(int, int), XYZ>();
            double inv = 1.0 / Math.Max(tol, 1e-9);

            // RefPoint thắng
            foreach (var p in refPts)
            {
                int ix = (int)Math.Round(p.X * inv);
                int iy = (int)Math.Round(p.Y * inv);
                map[(ix, iy)] = p;
            }

            // Thêm điểm tính toán nếu cell chưa có RefPoint
            foreach (var p in computed)
            {
                int ix = (int)Math.Round(p.X * inv);
                int iy = (int)Math.Round(p.Y * inv);
                var key = (ix, iy);
                if (!map.ContainsKey(key)) map[key] = p;
            }

            return map.Values.ToList();
        }

        // 7. Ép thêm các RefPoints vào
        private List<XYZ> ForceIncludeRefPoints(List<XYZ> allPts, List<XYZ> refPts, double tol)
        {
            double inv = 1.0 / Math.Max(tol, 1e-9);
            var map = new Dictionary<(int, int), XYZ>();

            foreach (var p in allPts)
            {
                int ix = (int)Math.Round(p.X * inv);
                int iy = (int)Math.Round(p.Y * inv);
                var key = (ix, iy);
                if (!map.ContainsKey(key)) map[key] = p;
            }

            foreach (var rp in refPts) // RefPoint overwrite nếu trùng cell
            {
                int ix = (int)Math.Round(rp.X * inv);
                int iy = (int)Math.Round(rp.Y * inv);
                map[(ix, iy)] = rp;
            }

            return map.Values.ToList();
        }

        // 8. Phương pháp EnforceMaxPointsWithRefPriority
        private List<XYZ> EnforceMaxPointsWithRefPriority(List<XYZ> allPts, List<XYZ> refPts, int max, double tol)
        {
            if (allPts.Count <= max) return allPts;

            double inv = 1.0 / Math.Max(tol, 1e-9);
            var refKeys = new HashSet<(int, int)>(
                refPts.Select(p => ((int)Math.Round(p.X * inv), (int)Math.Round(p.Y * inv)))
            );

            var refs = new List<XYZ>();
            var others = new List<XYZ>();
            foreach (var p in allPts)
            {
                var k = ((int)Math.Round(p.X * inv), (int)Math.Round(p.Y * inv));
                if (refKeys.Contains(k)) refs.Add(p); else others.Add(p);
            }

            if (refs.Count >= max) return refs.Take(max).ToList();

            int need = max - refs.Count;
            refs.AddRange(Downsample(others, Math.Max(need, 0)));
            return refs;
        }

        private List<XYZ> DeduplicateXY(IEnumerable<XYZ> pts, double tol)
        {
            var map = new Dictionary<(int, int), XYZ>();
            double inv = 1.0 / Math.Max(tol, 1e-9);

            foreach (var p in pts)
            {
                int ix = (int)Math.Round(p.X * inv);
                int iy = (int)Math.Round(p.Y * inv);
                var key = (ix, iy);
                if (!map.ContainsKey(key)) map[key] = p; // giữ phần tử đầu tiên
            }
            return map.Values.ToList();
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

        private double DistXY(XYZ a, XYZ b)
        { double dx = a.X - b.X, dy = a.Y - b.Y; return Math.Sqrt(dx * dx + dy * dy); }
    }
}
