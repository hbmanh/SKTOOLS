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

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // B1. Chọn Toposolid
                Reference topoRef = uidoc.Selection.PickObject(ObjectType.Element, new ToposolidFilter(), "Chọn Toposolid muốn chỉnh sửa");
                Toposolid originalTopo = doc.GetElement(topoRef) as Toposolid;
                if (originalTopo == null)
                {
                    TaskDialog.Show("Lỗi", "Không phải là Toposolid.");
                    return Result.Failed;
                }

                // B2. Gom toàn bộ FamilyInstance RefPoint
                var refPoints = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.Equals(RefPointFamilyName, StringComparison.InvariantCultureIgnoreCase))
                    .ToList();

                if (!refPoints.Any())
                {
                    TaskDialog.Show("Lỗi", "Không tìm thấy RefPoint nào.");
                    return Result.Failed;
                }

                // Biên dạng toposolid
                IList<XYZ> topoBoundary = GetTopoBoundary(originalTopo);

                List<XYZ> insidePoints = new List<XYZ>();
                List<FamilyInstance> outsideRefs = new List<FamilyInstance>();

                foreach (var fi in refPoints)
                {
                    LocationPoint loc = fi.Location as LocationPoint;
                    if (loc == null) continue;

                    XYZ p = loc.Point;
                    double z = p.Z;

                    Parameter elevParam = fi.LookupParameter("Default Elevation");
                    if (elevParam != null && elevParam.StorageType == StorageType.Double)
                    {
                        z = elevParam.AsDouble();
                    }

                    XYZ pt = new XYZ(p.X, p.Y, z);

                    if (IsPointInsidePolygon(pt, topoBoundary))
                        insidePoints.Add(pt);
                    else
                        outsideRefs.Add(fi);
                }

                // Với các điểm ngoài biên → hiệu chỉnh Z theo RefPoint gần nhất trong biên
                foreach (var fi in outsideRefs)
                {
                    LocationPoint loc = fi.Location as LocationPoint;
                    if (loc == null) continue;

                    XYZ p = loc.Point;
                    double z = p.Z;

                    Parameter elevParam = fi.LookupParameter("Default Elevation");
                    if (elevParam != null && elevParam.StorageType == StorageType.Double)
                    {
                        z = elevParam.AsDouble();
                    }

                    XYZ rawP = new XYZ(p.X, p.Y, z);

                    if (insidePoints.Count > 0)
                    {
                        XYZ nearest = FindNearestPoint(rawP, insidePoints);
                        XYZ newP = new XYZ(rawP.X, rawP.Y, nearest.Z);
                        insidePoints.Add(newP);
                    }
                }

                // Bổ sung điểm trên boundary với cao độ bằng nearest inside point
                AddBoundaryPointsWithNearestElevation(topoBoundary, insidePoints);

                if (insidePoints.Count == 0)
                {
                    TaskDialog.Show("Lỗi", "Không có điểm nào hợp lệ để tạo Toposolid.");
                    return Result.Failed;
                }

                // B3. Lấy type + level
                ElementId typeId = originalTopo.GetTypeId();
                Level level = GetNearestLevel(doc, originalTopo);
                if (level == null)
                {
                    TaskDialog.Show("Lỗi", "Không tìm thấy Level.");
                    return Result.Failed;
                }

                // B4. Tạo lại Toposolid
                using (Transaction t = new Transaction(doc, "Cập nhật Toposolid từ RefPoint"))
                {
                    t.Start();

                    IList<CurveLoop> boundaryLoops = GetToposolidBoundary(originalTopo);
                    Toposolid newTopo = Toposolid.Create(doc, boundaryLoops, insidePoints, typeId, level.Id);

                    // Xoá Toposolid cũ
                    doc.Delete(originalTopo.Id);

                    t.Commit();
                }

                TaskDialog.Show("Xong", $"Tạo Toposolid mới với {insidePoints.Count} điểm (bao gồm RefPoint và boundary).");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private class ToposolidFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) =>
                elem.Category != null &&
                elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Toposolid;

            public bool AllowReference(Reference reference, XYZ position) => true;
        }

        private IList<XYZ> GetTopoBoundary(Toposolid topo)
        {
            IList<XYZ> boundary = new List<XYZ>();

            Options opt = new Options
            {
                ComputeReferences = false,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = false
            };

            GeometryElement geomElem = topo.get_Geometry(opt);
            foreach (GeometryObject obj in geomElem)
            {
                if (obj is Solid solid && solid.Faces.Size > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace pf &&
                            pf.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ.Negate())) // mặt đáy
                        {
                            IList<CurveLoop> loops = face.GetEdgesAsCurveLoops();
                            if (loops.Count > 0)
                            {
                                // lấy loop lớn nhất (chu vi ngoài cùng)
                                CurveLoop outer = loops.OrderByDescending(l => l.GetExactLength()).First();

                                foreach (Curve c in outer)
                                {
                                    boundary.Add(c.GetEndPoint(0));
                                }
                                return boundary;
                            }
                        }
                    }
                }
            }

            return boundary;
        }

        private bool IsPointInsidePolygon(XYZ p, IList<XYZ> polygon)
        {
            int n = polygon.Count;
            bool inside = false;

            for (int i = 0, j = n - 1; i < n; j = i++)
            {
                if (((polygon[i].Y > p.Y) != (polygon[j].Y > p.Y)) &&
                     (p.X < (polygon[j].X - polygon[i].X) *
                      (p.Y - polygon[i].Y) /
                      (polygon[j].Y - polygon[i].Y) + polygon[i].X))
                {
                    inside = !inside;
                }
            }
            return inside;
        }

        private XYZ FindNearestPoint(XYZ target, List<XYZ> points)
        {
            XYZ nearest = points[0];
            double minDist = target.DistanceTo(nearest);

            foreach (var pt in points)
            {
                double d = target.DistanceTo(pt);
                if (d < minDist)
                {
                    minDist = d;
                    nearest = pt;
                }
            }
            return nearest;
        }

        private Level GetNearestLevel(Document doc, Toposolid topo)
        {
            BoundingBoxXYZ bbox = topo.get_BoundingBox(null);
            if (bbox == null) return null;

            double z = bbox.Min.Z;

            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => Math.Abs(l.Elevation - z))
                .ToList();

            return levels.FirstOrDefault();
        }

        private IList<CurveLoop> GetToposolidBoundary(Toposolid topo)
        {
            Options opt = new Options
            {
                ComputeReferences = false,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = false
            };

            GeometryElement geomElem = topo.get_Geometry(opt);
            foreach (GeometryObject obj in geomElem)
            {
                if (obj is Solid solid && solid.Faces.Size > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace pf &&
                            pf.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ.Negate()))
                        {
                            return face.GetEdgesAsCurveLoops();
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Thêm các điểm trên boundary (có thể chia nhỏ) với cao độ bằng nearest inside point
        /// </summary>
        private void AddBoundaryPointsWithNearestElevation(IList<XYZ> topoBoundary, List<XYZ> insidePoints)
        {
            // Chia nhỏ boundary để mép mượt hơn
            List<XYZ> denseBoundaryPts = new List<XYZ>();
            for (int i = 0; i < topoBoundary.Count; i++)
            {
                XYZ start = topoBoundary[i];
                XYZ end = topoBoundary[(i + 1) % topoBoundary.Count];
                double dist = start.DistanceTo(end);

                // Số đoạn chia, ví dụ mỗi 1m một điểm
                int segments = Math.Max(1, (int)(dist / 100.0));

                for (int s = 0; s <= segments; s++)
                {
                    double t = (double)s / segments;
                    XYZ pt = start + (end - start) * t;
                    denseBoundaryPts.Add(pt);
                }
            }

            // Gán cao độ cho các điểm boundary dựa trên nearest inside point
            foreach (var boundaryPt in denseBoundaryPts)
            {
                if (insidePoints.Count > 0)
                {
                    XYZ nearest = FindNearestPoint(boundaryPt, insidePoints);
                    XYZ newBoundaryPt = new XYZ(boundaryPt.X, boundaryPt.Y, nearest.Z);
                    insidePoints.Add(newBoundaryPt);
                }
            }
        }
    }
}
