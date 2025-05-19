using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;

namespace SKRevitAddins.PermissibleRangeFrame
{
    public class PermissibleRangeFrameRequestHandler : IExternalEventHandler
    {
        private PermissibleRangeFrameViewModel ViewModel { get; }
        private PermissibleRangeFrameRequest m_Request = new PermissibleRangeFrameRequest();

        public PermissibleRangeFrameRequest Request => m_Request;

        public PermissibleRangeFrameRequestHandler(PermissibleRangeFrameViewModel viewModel)
        {
            ViewModel = viewModel;
        }

        public void Execute(UIApplication uiapp)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.OK:
                        CreatePermissibleRange(uiapp, ViewModel);
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        public string GetName() => "PermissibleRangeFrameRequestHandler";

        #region Permissible Range

        public void CreatePermissibleRange(UIApplication uiapp, PermissibleRangeFrameViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Thu thập các MEPCurve trong View hiện tại
            var mepCurves = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfClass(typeof(MEPCurve))
                .Cast<MEPCurve>()
                .ToList();

            // Lấy các tham số từ ViewModel
            double x = viewModel.X, y = viewModel.Y, a = viewModel.A, b = viewModel.B, c = viewModel.C;
            var structuralFramings = viewModel.StructuralFramings;
            var sleeveSymbol = viewModel.SleeveSymbol;
            var intersectionData = viewModel.IntersectionData;
            var errorMessages = viewModel.ErrorMessages;
            var sleevePlacements = viewModel.SleevePlacements;
            var directShapes = viewModel.DirectShapes;

            // Xoá dữ liệu cũ
            intersectionData.Clear();
            errorMessages.Clear();
            sleevePlacements.Clear();
            directShapes.Clear();

            // 1) Tạo phạm vi cho phép xuyên (DirectShape) và lấy điểm giao
            ProcessIntersections(doc, structuralFramings, mepCurves, intersectionData, directShapes, x, y);

            // 2) Đặt Sleeve nếu đủ điều kiện
            PlaceSleeves(doc, sleeveSymbol, intersectionData, structuralFramings, sleevePlacements, errorMessages, directShapes, a, b, c);

            // 3) Tạo báo cáo lỗi (nếu cần)
            if (viewModel.CreateErrorSchedules && errorMessages.Any())
                CreateErrorSchedules(doc, errorMessages);

            // 4) Áp dụng filter cho DirectShape
            ApplyFilterToDirectShapes(doc, uidoc.ActiveView, directShapes);

            // 5) Cleanup
            using (Transaction subtx = new Transaction(doc, "Cleanup"))
            {
                subtx.Start();
                if (!viewModel.PlaceSleeves)
                {
                    var sleeves = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_PipeAccessory)
                        .OfClass(typeof(FamilyInstance))
                        .WhereElementIsNotElementType()
                        .ToList();
                    doc.Delete(sleeves.Select(s => s.Id).ToList());
                }
                if (!viewModel.PermissibleRange)
                {
                    doc.Delete(directShapes.Select(ds => ds.Id).ToList());
                }
                subtx.Commit();
            }
        }

        private void ProcessIntersections(
            Document doc,
            List<Element> structuralFramings,
            List<MEPCurve> mepCurves,
            Dictionary<(ElementId MEPCurveId, ElementId FrameId), List<XYZ>> intersectionData,
            List<DirectShape> directShapes,
            double x, double y)
        {
            using (Transaction trans = new Transaction(doc, "Create Permissible Range for Frame"))
            {
                trans.Start();
                foreach (var frame in structuralFramings)
                {
                    var frameObj = new PermissibleRangeFrameViewModel.FrameObj(frame);
                    var surroundingFaces = GetSurroundingFacesOfFrame(frameObj.FramingSolid);
                    foreach (Face face in surroundingFaces)
                    {
                        var directShape = CreateDirectShapeFromFrameFace(doc, frameObj, face, x, y);
                        if (directShape != null) directShapes.Add(directShape);

                        foreach (var pipeOrDuct in mepCurves)
                        {
                            var pipeOrDuctCurve = (pipeOrDuct.Location as LocationCurve)?.Curve;
                            if (pipeOrDuctCurve == null) continue;

                            if (face.Intersect(pipeOrDuctCurve, out IntersectionResultArray resultArray) != SetComparisonResult.Overlap)
                                continue;

                            var key = (MEPCurveId: pipeOrDuct.Id, FrameId: frame.Id);
                            if (!intersectionData.ContainsKey(key))
                                intersectionData[key] = new List<XYZ>();
                            foreach (IntersectionResult intersectionResult in resultArray)
                                intersectionData[key].Add(intersectionResult.XYZPoint);
                        }
                    }
                }
                trans.Commit();
            }
        }

        private DirectShape CreateDirectShapeFromFrameFace(
            Document doc,
            PermissibleRangeFrameViewModel.FrameObj frameObj,
            Face face,
            double x, double y)
        {
            if (face is not PlanarFace planarFace) return null;

            BoundingBoxUV bboxUV = planarFace.GetBoundingBox();
            UV minUV = bboxUV.Min;
            UV maxUV = bboxUV.Max;
            double frameHeight = frameObj.FramingHeight;
            double widthMargin = frameHeight * x;
            double heightMargin = frameHeight * y;
            UV adjustedMin = new(minUV.U + widthMargin, minUV.V + heightMargin);
            UV adjustedMax = new(maxUV.U - widthMargin, maxUV.V - heightMargin);

            if (adjustedMin.V >= adjustedMax.V)
            {
                adjustedMin = new UV(adjustedMin.U, minUV.V - heightMargin);
                adjustedMax = new UV(adjustedMax.U, maxUV.V + heightMargin);
            }
            if (adjustedMin.U >= adjustedMax.U)
            {
                adjustedMin = new UV(minUV.U + heightMargin, adjustedMin.V);
                adjustedMax = new UV(maxUV.U - heightMargin, adjustedMax.V);
            }

            List<Curve> profile = new();
            Func<XYZ, XYZ, bool> isValidCurve = (start, end) => start.DistanceTo(end) > doc.Application.ShortCurveTolerance;
            XYZ corner1 = face.Evaluate(adjustedMin);
            XYZ corner2 = face.Evaluate(new UV(adjustedMin.U, adjustedMax.V));
            XYZ corner3 = face.Evaluate(adjustedMax);
            XYZ corner4 = face.Evaluate(new UV(adjustedMax.U, adjustedMin.V));
            if (isValidCurve(corner1, corner2)) profile.Add(Line.CreateBound(corner1, corner2));
            if (isValidCurve(corner2, corner3)) profile.Add(Line.CreateBound(corner2, corner3));
            if (isValidCurve(corner3, corner4)) profile.Add(Line.CreateBound(corner3, corner4));
            if (isValidCurve(corner4, corner1)) profile.Add(Line.CreateBound(corner4, corner1));
            if (profile.Count < 3) return null;

            CurveLoop newCurveLoop = CurveLoop.Create(profile);
            XYZ extrusionDirection = planarFace.FaceNormal;
            double extrusionDistance = 10.0 / 304.8; // 10mm
            Solid directShapeSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                new List<CurveLoop> { newCurveLoop }, extrusionDirection, extrusionDistance);

            if (directShapeSolid != null && directShapeSolid.Volume > 0)
            {
                DirectShape directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                directShape.SetShape(new GeometryObject[] { directShapeSolid });
                directShape.Name = "Phạm vi cho phép xuyên dầm";
                return directShape;
            }
            return null;
        }

        private static List<Face> GetSurroundingFacesOfFrame(Solid solid)
        {
            var faces = solid.GetSolidVerticalFaces();
            var faceAreas = faces.Select(face =>
            {
                Mesh mesh = face.Triangulate();
                double area = 0;
                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    MeshTriangle triangle = mesh.get_Triangle(i);
                    XYZ p0 = triangle.get_Vertex(0);
                    XYZ p1 = triangle.get_Vertex(1);
                    XYZ p2 = triangle.get_Vertex(2);
                    area += 0.5 * ((p1 - p0).CrossProduct(p2 - p0)).GetLength();
                }
                return new { Face = face, Area = area };
            }).ToList();

            double minArea = faceAreas.Min(f => f.Area);
            return faceAreas.Where(f => f.Area > minArea).Select(f => f.Face).ToList();
        }

        /// <summary>
        /// Đặt Sleeves sau khi đã loại trừ các lỗi và sắp xếp điểm giao.
        /// </summary>
        private void PlaceSleeves(
            Document doc,
            FamilySymbol sleeveSymbol,
            Dictionary<(ElementId MEPCurveId, ElementId FrameId), List<XYZ>> intersectionData,
            List<Element> structuralFramings,
            Dictionary<ElementId, List<(XYZ, double)>> sleevePlacements,
            Dictionary<ElementId, HashSet<string>> errorMessages,
            List<DirectShape> directShapes,
            double a, double b, double c)
        {
            if (sleeveSymbol == null)
            {
                TaskDialog.Show("Thông báo", "Không tìm thấy Family Symbol cho Sleeve. Không thể tạo Sleeve.");
                return;
            }

            int successfulSleevesCount = 0;
            var potentialSleeves = new List<(ElementId MEPCurveId, XYZ Midpoint, double Diameter, double Length, XYZ Direction)>();

            foreach (var entry in intersectionData)
            {
                var mepCurveId = entry.Key.MEPCurveId;
                var frameId = entry.Key.FrameId;
                var pipeOrDuct = doc.GetElement(mepCurveId) as MEPCurve;
                var frame = structuralFramings.FirstOrDefault(f => f.Id == frameId);
                if (frame == null || pipeOrDuct == null) continue;

                var frameObj = new PermissibleRangeFrameViewModel.FrameObj(frame);
                var points = entry.Value;
                var pipeOrDuctCurve = (pipeOrDuct.Location as LocationCurve)?.Curve;
                if (pipeOrDuctCurve == null || points.Count < 2) continue;

                // Sắp xếp điểm giao theo trục ống
                points = points.OrderBy(pt => pipeOrDuctCurve.Project(pt).Parameter).ToList();

                for (int i = 0; i + 1 < points.Count; i += 2)
                {
                    XYZ point1 = points[i];
                    XYZ point2 = points[i + 1];
                    XYZ midpoint = (point1 + point2) / 2;
                    XYZ direction = (point2 - point1).Normalize();

                    double pipeDiameter = pipeOrDuct.LookupParameter("Diameter")?.AsDouble() ?? 0;
                    double sleeveDiameter = pipeDiameter + 50.0 / 304.8; // +50mm
                    double frameHeight = frameObj.FramingHeight;

                    // Kiểm tra OD > a (750mm) => Lỗi
                    if (sleeveDiameter > a / 304.8)
                    {
                        if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                            errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                        errorMessages[pipeOrDuct.Id].Add("OD > " + a + "mm");
                        continue;
                    }
                    // Kiểm tra OD > H/3 => Lỗi
                    if (sleeveDiameter > frameHeight * b)
                    {
                        if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                            errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                        errorMessages[pipeOrDuct.Id].Add("OD > H/3");
                        continue;
                    }

                    double sleeveLength = point1.DistanceTo(point2);

                    potentialSleeves.Add((mepCurveId, midpoint, sleeveDiameter, sleeveLength, direction));
                }
            }

            // Loại bỏ sleeves quá gần nhau, giữ lại 1 sleeve
            var sleevesToRemove = new HashSet<int>();
            for (int i = 0; i < potentialSleeves.Count; i++)
            {
                for (int j = i + 1; j < potentialSleeves.Count; j++)
                {
                    var s1 = potentialSleeves[i];
                    var s2 = potentialSleeves[j];

                    double minDistance = (s1.Diameter + s2.Diameter) * c;
                    double actualDistance = s1.Midpoint.DistanceTo(s2.Midpoint);

                    if (actualDistance < minDistance)
                    {
                        // Giữ lại sleeve có đường kính nhỏ hơn (nếu bằng thì random)
                        int removeIdx = s1.Diameter == s2.Diameter ?
                            (new Random()).Next(0, 2) == 0 ? i : j :
                            (s1.Diameter > s2.Diameter ? i : j);
                        sleevesToRemove.Add(removeIdx);

                        // Báo lỗi cho sleeve bị loại
                        var removeSleeve = potentialSleeves[removeIdx];
                        if (!errorMessages.ContainsKey(removeSleeve.MEPCurveId))
                            errorMessages[removeSleeve.MEPCurveId] = new HashSet<string>();
                        errorMessages[removeSleeve.MEPCurveId].Add("Khoảng cách giữa hai Sleeve < (OD1 + OD2)*2/3");
                    }
                }
            }

            // Giữ lại các sleeve hợp lệ
            var sleevesToKeep = potentialSleeves
                .Where((s, idx) => !sleevesToRemove.Contains(idx))
                .ToList();

            using (Transaction trans = new Transaction(doc, "Place Sleeve"))
            {
                trans.Start();
                foreach (var sleeveData in sleevesToKeep)
                {
                    var mepCurveId = sleeveData.MEPCurveId;
                    var midpoint = sleeveData.Midpoint;
                    var sleeveDiameter = sleeveData.Diameter;
                    var sleeveLength = sleeveData.Length;
                    var direction = sleeveData.Direction;
                    var pipeOrDuct = doc.GetElement(mepCurveId);

                    // Tạo FamilyInstance
                    FamilyInstance sleeveInstance = doc.Create.NewFamilyInstance(midpoint, sleeveSymbol, StructuralType.NonStructural);

                    // Xoay cho trùng hướng ống
                    Line rotationAxis = Line.CreateBound(midpoint, midpoint + XYZ.BasisZ);
                    double rotationAngle = XYZ.BasisX.AngleTo(direction) + Math.PI / 2;
                    ElementTransformUtils.RotateElement(doc, sleeveInstance.Id, rotationAxis, rotationAngle);

                    // Gán tham số
                    sleeveInstance.LookupParameter("L")?.Set(sleeveLength);
                    sleeveInstance.LookupParameter("OD")?.Set(sleeveDiameter);

                    // Lưu vị trí
                    if (!sleevePlacements.ContainsKey(mepCurveId))
                        sleevePlacements[mepCurveId] = new List<(XYZ, double)>();
                    sleevePlacements[mepCurveId].Add((midpoint, sleeveDiameter));

                    successfulSleevesCount++;

                    // Sửa: chỉ cần thuộc bất kỳ direct shape là hợp lệ
                    bool isWithinDirectShapes = directShapes.Any(ds => IsPointWithinDirectShape(midpoint, ds));
                    if (!isWithinDirectShapes)
                    {
                        if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                            errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                        errorMessages[pipeOrDuct.Id].Add("Sleeve nằm ngoài phạm vi cho phép xuyên dầm.");
                    }
                }
                trans.Commit();
            }

            TaskDialog.Show("Thông báo:", $"Đã đặt {successfulSleevesCount} Sleeves thỏa điều kiện. Vui lòng kiểm tra lại.");
        }

        private bool IsPointWithinDirectShape(XYZ point, DirectShape directShape)
        {
            var geometryElement = directShape.get_Geometry(new Options());
            foreach (GeometryObject geomObj in geometryElement)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        var result = face.Project(point);
                        if (result != null && result.Distance < 0.001)
                            return true;
                    }
                }
            }
            return false;
        }

        private void CreateErrorSchedules(Document doc, Dictionary<ElementId, HashSet<string>> errorMessages)
        {
            using (Transaction subtx = new Transaction(doc, "Create Report"))
            {
                subtx.Start();
                var schedules = new List<(string scheduleName, BuiltInCategory category)>
                {
                    ("PipeErrorSchedule", BuiltInCategory.OST_PipeCurves),
                    ("DuctErrorSchedule", BuiltInCategory.OST_DuctCurves)
                };
                foreach (var (scheduleName, category) in schedules)
                {
                    var existingSchedule = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSchedule))
                        .Cast<ViewSchedule>()
                        .FirstOrDefault(sch => sch.Name.Equals(scheduleName));
                    if (existingSchedule != null)
                        doc.Delete(existingSchedule.Id);

                    ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, new ElementId(category));
                    schedule.Name = scheduleName;

                    var schedulableFields = schedule.Definition.GetSchedulableFields();
                    var markField = schedulableFields.FirstOrDefault(sf => sf.GetName(doc) == "Mark");
                    var commentField = schedulableFields.FirstOrDefault(sf => sf.GetName(doc) == "Comments");
                    if (markField != null) schedule.Definition.AddField(markField);
                    if (commentField != null) schedule.Definition.AddField(commentField);

                    var elementsToClear = new FilteredElementCollector(doc)
                        .OfCategory(category)
                        .WhereElementIsNotElementType()
                        .ToList();
                    foreach (var element in elementsToClear)
                    {
                        element.LookupParameter("Mark")?.Set(string.Empty);
                        element.LookupParameter("Comments")?.Set(string.Empty);
                    }

                    int markIndex = 1;
                    foreach (var kvp in errorMessages)
                    {
                        Element element = doc.GetElement(kvp.Key);
                        if (element?.Category == null) continue;
                        if (element.Category.Id.IntegerValue != (int)category)
                            continue;
                        var markParam = element.LookupParameter("Mark");
                        if (markParam != null && kvp.Value.Any())
                        {
                            markParam.Set(markIndex.ToString());
                            markIndex++;
                        }
                        element.LookupParameter("Comments")?.Set(string.Join(", ", kvp.Value));
                    }
                }
                subtx.Commit();
            }
        }

        private void ApplyFilterToDirectShapes(Document doc, View view, List<DirectShape> directShapes)
        {
            using (Transaction trans = new Transaction(doc, "Apply DirectShape Filter"))
            {
                trans.Start();
                var parameterFilter = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .FirstOrDefault(f => f.Name == "DirectShape Filter");
                if (parameterFilter == null)
                    parameterFilter = ParameterFilterElement.Create(
                        doc, "DirectShape Filter", new List<ElementId> { new ElementId(BuiltInCategory.OST_GenericModel) });

                if (!view.IsFilterApplied(parameterFilter.Id))
                    view.AddFilter(parameterFilter.Id);

                ElementId solidFillPatternId = null;
                var fillPatternElements = new FilteredElementCollector(doc)
                    .OfClass(typeof(FillPatternElement))
                    .Cast<FillPatternElement>();
                foreach (var fp in fillPatternElements)
                {
                    if (fp.GetFillPattern().IsSolidFill)
                    {
                        solidFillPatternId = fp.Id;
                        break;
                    }
                }

                OverrideGraphicSettings overrideSettings = new();
                overrideSettings.SetSurfaceForegroundPatternColor(new Color(0, 255, 0));
                if (solidFillPatternId != null)
                    overrideSettings.SetSurfaceForegroundPatternId(solidFillPatternId);

                foreach (var ds in directShapes)
                {
                    view.SetElementOverrides(ds.Id, overrideSettings);
                }
                trans.Commit();
            }
        }
        #endregion
    }
}
