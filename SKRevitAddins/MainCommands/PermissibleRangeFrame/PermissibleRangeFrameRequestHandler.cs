#region ▬ using ▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;                      // Unit-utils + TUnionUtils
using SKRevitAddins.PermissibleRangeFrame;     // MEPCurveExtensions
#endregion

namespace SKRevitAddins.PermissibleRangeFrame
{
    /// <summary>
    /// Handler: sinh “phạm vi xuyên dầm” (1 DirectShape/dầm) & đặt Sleeve cho ống tròn.
    /// Điều kiện: (a) OD≤A, (b) OD≤B·H, (c) KC≥(OD1+OD2)·C, (d) chu vi Sleeve
    /// nằm trọn trong DirectShape gộp.
    /// </summary>
    public class PermissibleRangeFrameRequestHandler : IExternalEventHandler
    {
        #region ▶ Boilerplate
        private readonly PermissibleRangeFrameViewModel _vm;
        private readonly PermissibleRangeFrameRequest _req = new();
        public PermissibleRangeFrameRequest Request => _req;

        public PermissibleRangeFrameRequestHandler(PermissibleRangeFrameViewModel vm) => _vm = vm;
        public string GetName() => nameof(PermissibleRangeFrameRequestHandler);

        public void Execute(UIApplication uiapp)
        {
            try
            {
                if (Request.Take() == RequestId.OK)
                    CreatePermissibleRange(uiapp, _vm);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.ToString());
            }
        }
        #endregion

        #region ▶ Entry-point chính
        private void CreatePermissibleRange(UIApplication uiapp, PermissibleRangeFrameViewModel vm)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            /* 0. Thu thập ống tròn trong view */
            var meps = new FilteredElementCollector(doc, doc.ActiveView.Id)
                           .OfClass(typeof(MEPCurve))
                           .Cast<MEPCurve>()
                           .Where(m => m.LookupParameter("Diameter")?.AsDouble() > 0
                                    || m.LookupParameter("Radius")?.AsDouble() > 0)
                           .ToList();

            /* 1. Reset ViewModel */
            vm.IntersectionData.Clear();
            vm.ErrorMessages.Clear();
            vm.SleevePlacements.Clear();
            vm.DirectShapes.Clear();

            /* 2. Tạo DirectShape gộp & giao điểm */
            var frameDsMap = new Dictionary<ElementId, DirectShape>();

            ProcessIntersections(
                doc,
                vm.StructuralFramings,
                meps,
                vm.IntersectionData,
                vm.DirectShapes,
                frameDsMap,
                vm.X,
                vm.Y);

            /* 3. Đặt Sleeve */
            PlaceSleeves(
                doc,
                vm.SleeveSymbol,
                vm.IntersectionData,
                vm.StructuralFramings,
                vm.SleevePlacements,
                vm.ErrorMessages,
                frameDsMap,
                vm.A,
                vm.B,
                vm.C);

            /* 4. Schedule lỗi (nếu bật) */
            if (vm.CreateErrorSchedules && vm.ErrorMessages.Any())
                CreateErrorSchedules(doc, vm.ErrorMessages);

            /* 5. Tô màu DS */
            ApplyFilterToDirectShapes(doc, uidoc.ActiveView, vm.DirectShapes);

            /* 6. Cleanup tuỳ chọn */
            using (var tx = new Transaction(doc, "Cleanup"))
            {
                tx.Start();

                if (!vm.PlaceSleeves)
                {
                    var sleeves = new FilteredElementCollector(doc)
                                    .OfCategory(BuiltInCategory.OST_PipeAccessory)
                                    .WhereElementIsNotElementType()
                                    .ToElementIds();
                    doc.Delete(sleeves);
                }

                if (!vm.PermissibleRange)
                    doc.Delete(vm.DirectShapes.Select(ds => ds.Id).ToList());

                tx.Commit();
            }
        }
        #endregion

        #region ▶ 2A. Sinh DirectShape gộp & Intersection
        //  ───────────────────────────────────────────────────────────────
        //  ProcessIntersections – phiên bản union 2 khối thành 1
        //  ───────────────────────────────────────────────────────────────
        private void ProcessIntersections(
    Document doc,
    List<Element> frames,
    List<MEPCurve> meps,
    Dictionary<(ElementId mepId, ElementId frameId), List<XYZ>> mapIntersect,
    List<DirectShape> directShapes,
    Dictionary<ElementId, DirectShape> mapFrameDS,
    double xRatio,
    double yRatio)
        {
            using var tx = new Transaction(doc, "Create permissible range");
            tx.Start();

            foreach (var frm in frames)
            {
                var fObj = new PermissibleRangeFrameViewModel.FrameObj(frm);

                /* 1. Hai mặt diện tích lớn nhất */
                var facesTop2 = GetSurroundingFacesOfFrame(fObj.FramingSolid)
                                   .OrderByDescending(fc => (fc as PlanarFace)?.Area ?? 0)
                                   .Take(2);

                var solids = new List<Solid>();

                /* 2. Tạo Solid extrusion cho từng mặt (KHÔNG tạo DirectShape ở đây) */
                foreach (var face in facesTop2)
                {
                    var solid = CreateExtrusionSolid(face, fObj, xRatio, yRatio);
                    if (solid != null) solids.Add(solid);

                    /* Giao điểm ống–dầm */
                    foreach (var mep in meps)
                    {
                        var axis = (mep.Location as LocationCurve)?.Curve;
                        if (axis == null) continue;

                        if (face.Intersect(axis, out IntersectionResultArray res)
                                != SetComparisonResult.Overlap) continue;

                        var key = (mep.Id, frm.Id);
                        if (!mapIntersect.ContainsKey(key))
                            mapIntersect[key] = new List<XYZ>();

                        foreach (IntersectionResult ir in res)
                            mapIntersect[key].Add(ir.XYZPoint);
                    }
                }

                if (solids.Count == 0) continue;

                /* 3. UNION 2 Solid → 1 Solid */
                Solid union = TUnionUtils.UnionSolids(solids);

                /* 4. Chỉ sinh 1 DirectShape duy nhất cho dầm */
                var ds = DirectShape.CreateElement(
                            doc, new ElementId(BuiltInCategory.OST_GenericModel));
                ds.Name = "Phạm vi cho phép xuyên dầm";
                ds.SetShape(new GeometryObject[] { union });

                directShapes.Add(ds);
                mapFrameDS[frm.Id] = ds;          // để bước PlaceSleeves tra nhanh
            }
            tx.Commit();
        }



        /// <summary>Extrusion Solid 10 mm vào trong dầm dựa trên PlanarFace.</summary>
        /// <summary>
        /// Tạo Solid extrude: lùi ra ngoài 10 mm rồi đùn vào trong
        /// (bề rộng dầm + 20 mm) theo hướng −normal.
        /// </summary>
        private Solid? CreateExtrusionSolid(
            Face face,
            PermissibleRangeFrameViewModel.FrameObj frmObj,
            double xRatio,
            double yRatio)
        {
            if (face is not PlanarFace pf) return null;

            /* 1. Thu hẹp khung UV theo xRatio, yRatio */
            var bb = pf.GetBoundingBox();
            double h = frmObj.FramingHeight;
            var minAdj = new UV(bb.Min.U + h * xRatio, bb.Min.V + h * yRatio);
            var maxAdj = new UV(bb.Max.U - h * xRatio, bb.Max.V - h * yRatio);

            if (minAdj.U >= maxAdj.U) { minAdj = new UV(bb.Min.U, minAdj.V); maxAdj = new UV(bb.Max.U, maxAdj.V); }
            if (minAdj.V >= maxAdj.V) { minAdj = new UV(minAdj.U, bb.Min.V); maxAdj = new UV(maxAdj.U, bb.Max.V); }

            /* 2. Tạo profile (4 đoạn) */
            XYZ p1 = pf.Evaluate(minAdj);
            XYZ p2 = pf.Evaluate(new UV(minAdj.U, maxAdj.V));
            XYZ p3 = pf.Evaluate(maxAdj);
            XYZ p4 = pf.Evaluate(new UV(maxAdj.U, minAdj.V));

            double tol = frmObj.FramingObj.Document.Application.ShortCurveTolerance;
            var profile = new List<Curve>();
            if (p1.DistanceTo(p2) > tol) profile.Add(Line.CreateBound(p1, p2));
            if (p2.DistanceTo(p3) > tol) profile.Add(Line.CreateBound(p2, p3));
            if (p3.DistanceTo(p4) > tol) profile.Add(Line.CreateBound(p3, p4));
            if (p4.DistanceTo(p1) > tol) profile.Add(Line.CreateBound(p4, p1));
            if (profile.Count < 3) return null;

            /* 3. Offset profile ra ngoài 10 mm theo +normal */
            double extra = 10.0.ToInternalUnits();
            XYZ vec = pf.FaceNormal.Multiply(extra);               // +normal 10 mm

            var moved = profile.Select(c =>
                        (c is Line ln)
                            ? (Curve)Line.CreateBound(ln.GetEndPoint(0) + vec,
                                                        ln.GetEndPoint(1) + vec)
                            : c.CreateTransformed(Transform.CreateTranslation(vec)))
                        .ToList();

            var loopMoved = CurveLoop.Create(moved);

            /* 4. Tính chiều sâu: bề rộng dầm + 20 mm */
            BoundingBoxXYZ bbFrm = frmObj.FramingSolid.GetBoundingBox();
            double width = Math.Abs((bbFrm.Max - bbFrm.Min).DotProduct(pf.FaceNormal));
            double depth = width + extra * 2;                          // +20 mm

            /* 5. Đùn vào trong theo –normal */
            return GeometryCreationUtilities.CreateExtrusionGeometry(
                       new[] { loopMoved }, -pf.FaceNormal, depth);
        }

        private static List<Face> GetSurroundingFacesOfFrame(Solid solid)
        {
            var faces = solid.GetSolidVerticalFaces();

            var infos = faces.Select(f =>
            {
                var mesh = f.Triangulate();
                double area = 0;
                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    var tri = mesh.get_Triangle(i);
                    area += 0.5 * (tri.get_Vertex(1) - tri.get_Vertex(0))
                                   .CrossProduct(tri.get_Vertex(2) - tri.get_Vertex(0))
                                   .GetLength();
                }
                return (Face: f, Area: area);
            })
            .OrderByDescending(x => x.Area)
            .ToList();

            double minArea = infos.Last().Area;
            return infos.Where(x => x.Area > minArea).Select(x => x.Face).ToList();
        }
        #endregion

        #region ▶ 2B. Đặt Sleeve (a, b, c + phạm vi)
        private void PlaceSleeves(
            Document doc,
            FamilySymbol sleeveSym,
            Dictionary<(ElementId mepId, ElementId frameId), List<XYZ>> intersectMap,
            List<Element> frames,
            Dictionary<ElementId, List<(XYZ, double)>> sleevePlaced,
            Dictionary<ElementId, HashSet<string>> errs,
            Dictionary<ElementId, DirectShape> frameDsMap,
            double a_mm,
            double b_ratio,
            double c_ratio)
        {
            if (sleeveSym == null)
            {
                TaskDialog.Show("Thông báo", "Không tìm thấy Family Symbol cho Sleeve.");
                return;
            }

            /* 1. Gom ứng viên (điều kiện a & b) */
            var cand = new List<SleeveCand>();

            foreach (var kv in intersectMap)
            {
                var mep = doc.GetElement(kv.Key.mepId) as MEPCurve;
                var frm = frames.FirstOrDefault(f => f.Id == kv.Key.frameId);
                if (mep == null || frm == null) continue;

                double odPipe = mep.GetOuterDiameter();
                if (odPipe <= 0) continue;

                var frmObj = new PermissibleRangeFrameViewModel.FrameObj(frm);
                double hFrm = frmObj.FramingHeight;

                var axis = (mep.Location as LocationCurve)?.Curve;
                if (axis == null || kv.Value.Count < 2) continue;

                var pts = kv.Value.Distinct()
                                   .OrderBy(pt => axis.Project(pt).Parameter)
                                   .ToList();

                for (int i = 0; i + 1 < pts.Count; i += 2)
                {
                    XYZ p1 = pts[i];
                    XYZ p2 = pts[i + 1];
                    var mid = (p1 + p2) / 2;
                    var dir = (p2 - p1).Normalize();

                    double odSleeve = odPipe + 50.0.ToInternalUnits();
                    double lenSleeve = p1.DistanceTo(p2);

                    if (odSleeve > a_mm.ToInternalUnits())
                    { AddErr(errs, mep.Id, $"OD > {a_mm} mm"); continue; }

                    if (odSleeve > hFrm * b_ratio)
                    { AddErr(errs, mep.Id, $"OD > H·{b_ratio:0.##}"); continue; }

                    cand.Add(new SleeveCand
                    {
                        MepId = mep.Id,
                        FrameId = frm.Id,
                        Mid = mid,
                        Dir = dir,
                        OD = odSleeve,
                        Len = lenSleeve
                    });
                }
            }

            /* 2. Điều kiện c – khoảng cách */
            foreach (var grp in cand.GroupBy(c => c.FrameId))
            {
                var list = grp.ToList();
                for (int i = 0; i < list.Count; ++i)
                    for (int j = i + 1; j < list.Count; ++j)
                    {
                        double minD = (list[i].OD + list[j].OD) * c_ratio;
                        double distT = list[i].Mid.DistanceTo(list[j].Mid);

                        if (distT < minD)
                        {
                            string msg = $"Khoảng cách < (OD₁+OD₂)*{c_ratio:0.##}";
                            AddErr(errs, list[i].MepId, msg);
                            AddErr(errs, list[j].MepId, msg);
                        }
                    }
            }

            /* 3. Kiểm chu vi nằm trong DS gộp & tạo sleeve */
            int ok = 0;
            using var tx = new Transaction(doc, "Place sleeves");
            tx.Start();
            if (!sleeveSym.IsActive) sleeveSym.Activate();

            foreach (var s in cand)
            {
                if (!frameDsMap.TryGetValue(s.FrameId, out var ds))
                {
                    AddErr(errs, s.MepId, "Không tìm thấy DS phạm vi");
                    continue;
                }

                bool inside = PermissibleRangeHelpers.CircleInsideShapes(
                                  new[] { ds }, s.Mid, s.Dir, s.OD / 2, 12);

                if (!inside)
                {
                    AddErr(errs, s.MepId, "Chu vi Sleeve vượt khỏi phạm vi");
                    continue;
                }

                var fi = doc.Create.NewFamilyInstance(
                             s.Mid, sleeveSym, StructuralType.NonStructural);

                var axisRot = Line.CreateBound(s.Mid, s.Mid + XYZ.BasisZ);
                double ang = XYZ.BasisX.AngleTo(s.Dir) + Math.PI / 2;
                ElementTransformUtils.RotateElement(doc, fi.Id, axisRot, ang);

                fi.LookupParameter("L")?.Set(s.Len);
                fi.LookupParameter("OD")?.Set(s.OD);

                if (!sleevePlaced.ContainsKey(s.MepId))
                    sleevePlaced[s.MepId] = new List<(XYZ, double)>();
                sleevePlaced[s.MepId].Add((s.Mid, s.OD));

                ok++;
            }
            tx.Commit();

            TaskDialog.Show("Thông báo", $"Đã đặt {ok} Sleeve hợp lệ (có thể kèm cảnh báo).");
        }

        private struct SleeveCand
        {
            public ElementId MepId;
            public ElementId FrameId;
            public XYZ Mid;
            public XYZ Dir;
            public double OD;
            public double Len;
        }

        private static void AddErr(Dictionary<ElementId, HashSet<string>> map, ElementId id, string msg)
        {
            if (!map.TryGetValue(id, out var set))
                map[id] = set = new HashSet<string>();
            set.Add(msg);
        }
        #endregion

        #region ▶ 3. Schedule lỗi & Tô màu DirectShape
        private void CreateErrorSchedules(Document doc, Dictionary<ElementId, HashSet<string>> err)
        {
            using var tx = new Transaction(doc, "Create error schedules");
            tx.Start();

            var targets = new[]
            {
                ("PipeErrorSchedule", BuiltInCategory.OST_PipeCurves),
                ("DuctErrorSchedule", BuiltInCategory.OST_DuctCurves)
            };

            foreach (var t in targets)
            {
                var exist = new FilteredElementCollector(doc)
                                .OfClass(typeof(ViewSchedule))
                                .Cast<ViewSchedule>()
                                .FirstOrDefault(v => v.Name == t.Item1);
                if (exist != null) doc.Delete(exist.Id);

                var vs = ViewSchedule.CreateSchedule(doc, new ElementId(t.Item2));
                vs.Name = t.Item1;

                var fields = vs.Definition.GetSchedulableFields();
                void AddField(string n)
                {
                    var f = fields.FirstOrDefault(ff => ff.GetName(doc) == n);
                    if (f != null) vs.Definition.AddField(f);
                }
                AddField("Mark");
                AddField("Comments");

                /* clear cũ */
                foreach (var el in new FilteredElementCollector(doc)
                                        .OfCategory(t.Item2)
                                        .WhereElementIsNotElementType())
                {
                    el.LookupParameter("Mark")?.Set(string.Empty);
                    el.LookupParameter("Comments")?.Set(string.Empty);
                }

                int idx = 1;
                foreach (var kv in err)
                {
                    var el = doc.GetElement(kv.Key);
                    if (el?.Category?.Id.IntegerValue != (int)t.Item2) continue;

                    if (kv.Value.Any())
                        el.LookupParameter("Mark")?.Set(idx++.ToString());

                    el.LookupParameter("Comments")
                       ?.Set(string.Join(", ", kv.Value));
                }
            }
            tx.Commit();
        }

        private void ApplyFilterToDirectShapes(Document doc, View view, List<DirectShape> dss)
        {
            using var tx = new Transaction(doc, "Apply DirectShape filter");
            tx.Start();

            var filter = new FilteredElementCollector(doc)
                             .OfClass(typeof(ParameterFilterElement))
                             .Cast<ParameterFilterElement>()
                             .FirstOrDefault(f => f.Name == "DirectShape Filter")
                          ?? ParameterFilterElement.Create(
                                 doc, "DirectShape Filter",
                                 new[] { new ElementId(BuiltInCategory.OST_GenericModel) });

            if (!view.IsFilterApplied(filter.Id))
                view.AddFilter(filter.Id);

            var fpSolid = new FilteredElementCollector(doc)
                            .OfClass(typeof(FillPatternElement))
                            .Cast<FillPatternElement>()
                            .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill)
                           ?.Id;

            var ogs = new OverrideGraphicSettings();
            ogs.SetSurfaceForegroundPatternColor(new Color(0, 255, 0));
            if (fpSolid != null)
                ogs.SetSurfaceForegroundPatternId(fpSolid);

            foreach (var ds in dss)
                view.SetElementOverrides(ds.Id, ogs);

            tx.Commit();
        }
        #endregion
    }

    #region ▶ Helper: kiểm chu vi Sleeve
    internal static class PermissibleRangeHelpers
    {
        public static bool CircleInsideShapes(
            IEnumerable<DirectShape> dss,
            XYZ center,
            XYZ dir,
            double radius,
            int nSeg = 12,
            double tol = 1e-4)
        {
            XYZ v1 = dir.IsAlmostEqualTo(XYZ.BasisZ)
                   ? XYZ.BasisX
                   : XYZ.BasisZ.CrossProduct(dir).Normalize();
            XYZ v2 = dir.CrossProduct(v1).Normalize();

            var solids = dss.SelectMany(ds =>
                             ds.get_Geometry(new Options())
                               .OfType<Solid>()
                               .Where(s => s.Volume > 0))
                            .ToList();
            if (!solids.Any()) return false;

            for (int i = 0; i < nSeg; ++i)
            {
                double ang = 2 * Math.PI * i / nSeg;
                XYZ pt = center + radius * (Math.Cos(ang) * v1 + Math.Sin(ang) * v2);

                if (!solids.Any(s => s.IsPointInside(pt, tol)))
                    return false;
            }
            return true;
        }
    }
    #endregion
}
