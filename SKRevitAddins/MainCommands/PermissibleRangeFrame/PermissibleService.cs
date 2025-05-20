using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;

namespace SKRevitAddins.PermissibleRangeFrame
{
    /// <summary>
    /// Toàn bộ thuật toán (Parallel + Cache + SolidCurveIntersection).
    /// </summary>
    public sealed class PermissibleService : IPermissibleService
    {
        public async Task<PermissibleResult> RunAsync(
            UIApplication uiApp,
            PermissibleArgs a,
            IProgress<int> prog,
            CancellationToken tk)
            => await Task.Run(() =>
            {
                Document doc = uiApp.ActiveUIDocument.Document;
                var directShapes = new ConcurrentBag<DirectShape>();
                var sleeves = new ConcurrentBag<SleeveInfo>();
                var errors = new ConcurrentDictionary<ElementId, HashSet<string>>();

                int done = 0;
                var solidCache = a.Frames.ToDictionary(f => f.Id,
                    f => new PermissibleRangeFrameViewModel.FrameObj(f).FramingSolid);

                Parallel.ForEach(a.Frames, new ParallelOptions
                {
                    CancellationToken = tk,
                    MaxDegreeOfParallelism = Environment.ProcessorCount
                }, frame =>
                {
                    Solid solid = solidCache[frame.Id];

                    foreach (var mep in a.MepCurves)
                    {
                        if (tk.IsCancellationRequested) break;

                        Curve curve = (mep.Location as LocationCurve)?.Curve;
                        if (curve == null) continue;

                        var sci = BooleanOperationsUtils.SolidCurveIntersection(solid, curve, out _);
                        if (sci == null || sci.SegmentCount == 0) continue;

                        List<XYZ> pts = sci.GetCurveSegments()
                            .SelectMany(s => new[] { s.GetEndPoint(0), s.GetEndPoint(1) })
                            .OrderBy(pt => curve.Project(pt).Parameter)
                            .ToList();
                        if (pts.Count < 2) continue;

                        // chỉ một DirectShape cho mỗi frame
                        directShapes.Add(CreateDS(doc, solid, a.X, a.Y));

                        for (int i = 0; i + 1 < pts.Count; i += 2)
                        {
                            XYZ mid = (pts[i] + pts[i + 1]) * .5;
                            XYZ dir = (pts[i + 1] - pts[i]).Normalize();

                            double od = (mep.LookupParameter("Diameter")?.AsDouble() ?? 0) + 50.0.Feet();
                            double h = solid.GetSolidHeight();

                            if (od > a.Amm.Feet())
                            { errors.AddErr(mep.Id, $"OD>{a.Amm}mm"); continue; }

                            if (od > h * a.B)
                            { errors.AddErr(mep.Id, "OD>H/3"); continue; }

                            sleeves.Add(new SleeveInfo(mep.Id, mid, od,
                                            pts[i].DistanceTo(pts[i + 1]), dir));
                        }
                    }
                    prog?.Report(Interlocked.Increment(ref done) * 100 / a.Frames.Count);
                });

                // loại sleeve quá gần nhau
                var list = sleeves.ToList();
                var remove = new HashSet<int>();

                for (int i = 0; i < list.Count; i++)
                    for (int j = i + 1; j < list.Count; j++)
                    {
                        if (list[i].Mid.DistanceTo(list[j].Mid) <
                            (list[i].Diameter + list[j].Diameter) * a.C)
                        {
                            int rem = list[i].Diameter == list[j].Diameter
                                      ? (Random.Shared.Next(2) == 0 ? i : j)
                                      : (list[i].Diameter > list[j].Diameter ? i : j);
                            remove.Add(rem);
                            errors.AddErr(list[rem].CurveId,
                                 "Khoảng cách giữa 2 Sleeve < (OD1+OD2)*2/3");
                        }
                    }

                return new PermissibleResult(
                    directShapes.Distinct().ToList(),
                    list.Where((s, idx) => !remove.Contains(idx)).ToList(),
                    errors);
            }, tk);

        // ► helper DirectShape nhanh dạng hộp
        private static DirectShape CreateDS(Document doc, Solid solid, double x, double y)
        {
            BoundingBoxXYZ bb = solid.GetBoundingBox();
            double dz = bb.Max.Z - bb.Min.Z;

            var box = new BoundingBoxXYZ
            {
                Min = bb.Min + new XYZ(dz * x, 0, dz * y),
                Max = bb.Max - new XYZ(dz * x, 0, dz * y)
            };

            var ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
            var loop = CurveLoop.CreateRectangle(box);
            ds.SetShape(new[] {
                GeometryCreationUtilities.CreateExtrusionGeometry(
                    new() { loop }, XYZ.BasisY, bb.Max.Y - bb.Min.Y)
            });
            ds.Name = "PermissibleRange(DS)";
            return ds;
        }
    }

    internal static class DicExt
    {
        public static void AddErr(this ConcurrentDictionary<ElementId, HashSet<string>> dic,
                                  ElementId id, string msg)
        {
            dic.AddOrUpdate(id, _ => new() { msg },
                (_, set) => { set.Add(msg); return set; });
        }
    }
}
