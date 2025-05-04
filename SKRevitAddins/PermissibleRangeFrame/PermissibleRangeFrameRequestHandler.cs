using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;
using SKRevitAddins.ViewModel;
using static SKRevitAddins.ViewModel.PermissibleRangeFrameViewModel;
using Document = Autodesk.Revit.DB.Document;
using UnitUtils = SKRevitAddins.Utils.UnitUtils;
using View = Autodesk.Revit.DB.View;

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
                    case RequestId.None:
                        break;
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
            double x = viewModel.X;
            double y = viewModel.Y;
            double a = viewModel.A;
            double b = viewModel.B;
            double c = viewModel.C;

            // Lấy các danh sách, dictionary từ ViewModel
            var structuralFramings = viewModel.StructuralFramings;
            var sleeveSymbol = viewModel.SleeveSymbol;
            var intersectionData = viewModel.IntersectionData;
            var errorMessages = viewModel.ErrorMessages;
            var sleevePlacements = viewModel.SleevePlacements;
            var directShapes = viewModel.DirectShapes;

            // 1) Tạo phạm vi cho phép xuyên (DirectShape) và lấy điểm giao
            ProcessIntersections(doc, structuralFramings, mepCurves, intersectionData, directShapes, x, y);

            // 2) Đặt Sleeve nếu đủ điều kiện
            PlaceSleeves(doc, sleeveSymbol, intersectionData, structuralFramings, sleevePlacements, errorMessages, directShapes, a, b, c);

            // 3) Tạo báo cáo lỗi (nếu cần)
            if (viewModel.CreateErrorSchedules && errorMessages.Any())
                CreateErrorSchedules(doc, errorMessages);

            // 4) Áp dụng filter cho DirectShape
            ApplyFilterToDirectShapes(doc, uidoc.ActiveView, directShapes);

            // 5) Tuỳ chọn: nếu người dùng bỏ chọn PlaceSleeves hoặc PermissibleRange, thì xoá
            using (Transaction subtx = new Transaction(doc, "Cleanup"))
            {
                subtx.Start();

                // Nếu không muốn tạo sleeves, xoá tất cả FamilyInstance “Sleeve”
                if (!viewModel.PlaceSleeves)
                {
                    var sleeves = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_PipeAccessory)
                        .OfClass(typeof(FamilyInstance))
                        .WhereElementIsNotElementType()
                        .ToList();
                    var sleeveIds = sleeves.Select(s => s.Id).ToList();
                    doc.Delete(sleeveIds);
                }

                // Nếu không muốn tạo phạm vi, xoá tất cả DirectShape
                if (!viewModel.PermissibleRange)
                {
                    var directShapeIds = directShapes.Select(ds => ds.Id).ToList();
                    doc.Delete(directShapeIds);
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
                    var frameObj = new FrameObj(frame);

                    // Lấy các mặt xung quanh dầm
                    var surroundingFaces = GetSurroundingFacesOfFrame(frameObj.FramingSolid);

                    foreach (Face face in surroundingFaces)
                    {
                        // Tạo DirectShape biểu diễn vùng cho phép xuyên
                        var directShape = CreateDirectShapeFromFrameFace(doc, frameObj, face, x, y);
                        if (directShape != null)
                            directShapes.Add(directShape);

                        // Kiểm tra giao cắt giữa face và các MEPCurve
                        foreach (var pipeOrDuct in mepCurves)
                        {
                            var pipeOrDuctCurve = (pipeOrDuct.Location as LocationCurve)?.Curve;
                            if (pipeOrDuctCurve == null) continue;

                            if (face.Intersect(pipeOrDuctCurve, out IntersectionResultArray resultArray) != SetComparisonResult.Overlap)
                                continue;

                            // Thêm giao điểm vào dictionary
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
            FrameObj frameObj,
            Face face,
            double x,
            double y)
        {
            if (!(face is PlanarFace planarFace))
                return null;

            BoundingBoxUV bboxUV = planarFace.GetBoundingBox();
            UV minUV = bboxUV.Min;
            UV maxUV = bboxUV.Max;

            double frameHeight = frameObj.FramingHeight;
            double widthMargin = frameHeight * x;
            double heightMargin = frameHeight * y;

            // Tính toán toạ độ UV sau khi trừ margin
            UV adjustedMin = new UV(minUV.U + widthMargin, minUV.V + heightMargin);
            UV adjustedMax = new UV(maxUV.U - widthMargin, maxUV.V - heightMargin);

            // Đảm bảo adjustedMin < adjustedMax
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

            // Tạo profile (4 đường line) nếu valid
            List<Curve> profile = new List<Curve>();
            Func<XYZ, XYZ, bool> isValidCurve = (start, end) => start.DistanceTo(end) > doc.Application.ShortCurveTolerance;

            XYZ corner1 = face.Evaluate(adjustedMin);
            XYZ corner2 = face.Evaluate(new UV(adjustedMin.U, adjustedMax.V));
            XYZ corner3 = face.Evaluate(adjustedMax);
            XYZ corner4 = face.Evaluate(new UV(adjustedMax.U, adjustedMin.V));

            if (isValidCurve(corner1, corner2)) profile.Add(Line.CreateBound(corner1, corner2));
            if (isValidCurve(corner2, corner3)) profile.Add(Line.CreateBound(corner2, corner3));
            if (isValidCurve(corner3, corner4)) profile.Add(Line.CreateBound(corner3, corner4));
            if (isValidCurve(corner4, corner1)) profile.Add(Line.CreateBound(corner4, corner1));

            if (profile.Count < 3)
                return null;

            // Extrude để tạo Solid
            CurveLoop newCurveLoop = CurveLoop.Create(profile);
            XYZ extrusionDirection = planarFace.FaceNormal;
            double extrusionDistance = 10.0 / 304.8; // 10mm

            Solid directShapeSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                new List<CurveLoop> { newCurveLoop },
                extrusionDirection,
                extrusionDistance);

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
            // Lấy các mặt đứng xung quanh dầm
            var faces = solid.GetSolidVerticalFaces();

            // Tính diện tích từng face, loại bỏ face có diện tích min
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
            var remainingFaces = faceAreas
                .Where(f => f.Area > minArea)
                .Select(f => f.Face)
                .ToList();

            return remainingFaces;
        }

        /// <summary>
        /// Hàm PlaceSleeves: Kiểm tra điều kiện, đặt Sleeves. Nếu có lỗi, thêm vào errorMessages
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
            int successfulSleevesCount = 0;
            var potentialSleeves = new List<(ElementId MEPCurveId, XYZ Midpoint, double Diameter, double Length, XYZ Direction)>();

            // 1) Tạo list sleeves tiềm năng
            foreach (var entry in intersectionData)
            {
                var mepCurveId = entry.Key.MEPCurveId;
                var frameId = entry.Key.FrameId;

                var pipeOrDuct = doc.GetElement(mepCurveId);
                var frame = structuralFramings.FirstOrDefault(f => f.Id == frameId);
                if (frame == null) continue;

                var frameObj = new FrameObj(frame);
                var points = entry.Value;

                // Giả sử 2 điểm 1 cặp => midpoint
                for (int i = 0; i + 1 < points.Count; i += 2)
                {
                    XYZ point1 = points[i];
                    XYZ point2 = points[i + 1];
                    XYZ midpoint = (point1 + point2) / 2;
                    XYZ direction = (point2 - point1).Normalize();

                    double pipeDiameter = pipeOrDuct.LookupParameter("Diameter")?.AsDouble() ?? 0;
                    double sleeveDiameter = pipeDiameter + UnitUtils.MmToFeet(50); // +50mm
                    double frameHeight = frameObj.FramingHeight;

                    // Kiểm tra OD > 750mm => Lỗi
                    if (sleeveDiameter > UnitUtils.MmToFeet(a))
                    {
                        if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                            errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                        errorMessages[pipeOrDuct.Id].Add("OD > 750mm");
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

                    // Lưu lại sleeve tiềm năng
                    potentialSleeves.Add((mepCurveId, midpoint, sleeveDiameter, sleeveLength, direction));
                }
            }

            // 2) Loại bỏ các sleeves quá gần nhau
            var sleevesToRemove = new HashSet<int>();
            for (int i = 0; i < potentialSleeves.Count; i++)
            {
                for (int j = i + 1; j < potentialSleeves.Count; j++)
                {
                    var sleeve1 = potentialSleeves[i];
                    var sleeve2 = potentialSleeves[j];

                    double minDistance = (sleeve1.Diameter + sleeve2.Diameter) * c; // (OD1+OD2)*2/3
                    double actualDistance = sleeve1.Midpoint.DistanceTo(sleeve2.Midpoint);

                    if (actualDistance < minDistance)
                    {
                        sleevesToRemove.Add(i);
                        sleevesToRemove.Add(j);

                        // Thêm lỗi
                        if (!errorMessages.ContainsKey(sleeve1.MEPCurveId))
                            errorMessages[sleeve1.MEPCurveId] = new HashSet<string>();
                        errorMessages[sleeve1.MEPCurveId].Add("Khoảng cách giữa hai Sleeve < (OD1 + OD2)*2/3");

                        if (!errorMessages.ContainsKey(sleeve2.MEPCurveId))
                            errorMessages[sleeve2.MEPCurveId] = new HashSet<string>();
                        errorMessages[sleeve2.MEPCurveId].Add("Khoảng cách giữa hai Sleeve < (OD1 + OD2)*2/3");
                    }
                }
            }

            // Loại bỏ các sleeves vi phạm
            potentialSleeves = potentialSleeves
                .Where((s, index) => !sleevesToRemove.Contains(index))
                .ToList();

            // 3) Tạo FamilyInstance Sleeves
            using (Transaction trans = new Transaction(doc, "Place Sleeve"))
            {
                trans.Start();
                foreach (var sleeveData in potentialSleeves)
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

                    // Kiểm tra xem sleeve có nằm trong DirectShape hay không
                    bool isWithinDirectShapes = directShapes.All(ds => IsPointWithinDirectShape(midpoint, ds));
                    if (!isWithinDirectShapes)
                    {
                        // Thêm lỗi
                        if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                            errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                        errorMessages[pipeOrDuct.Id].Add("Sleeve nằm ngoài phạm vi cho phép xuyên dầm.");
                    }
                }
                trans.Commit();
            }

            // Thông báo số sleeves đặt thành công
            TaskDialog.Show("Thông báo:", $"Đã đặt {successfulSleevesCount} Sleeves thỏa điều kiện. Vui lòng kiểm tra lại.");
        }

        private bool IsPointWithinDirectShape(XYZ point, DirectShape directShape)
        {
            // Kiểm tra xem point có nằm trong solid của DirectShape không
            GeometryElement geometryElement = directShape.get_Geometry(new Options());
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
                    // Xoá schedule cũ nếu trùng tên
                    var existingSchedule = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSchedule))
                        .Cast<ViewSchedule>()
                        .FirstOrDefault(sch => sch.Name.Equals(scheduleName));

                    if (existingSchedule != null)
                        doc.Delete(existingSchedule.Id);

                    // Tạo schedule mới
                    ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, new ElementId(category));
                    schedule.Name = scheduleName;

                    // Thêm cột Mark, Comments
                    var schedulableFields = schedule.Definition.GetSchedulableFields();
                    var markField = schedulableFields.FirstOrDefault(sf => sf.GetName(doc) == "Mark");
                    var commentField = schedulableFields.FirstOrDefault(sf => sf.GetName(doc) == "Comments");

                    ScheduleField markScheduleField = markField != null
                        ? schedule.Definition.AddField(markField)
                        : null;
                    ScheduleField commentScheduleField = commentField != null
                        ? schedule.Definition.AddField(commentField)
                        : null;

                    // Clear cột Mark, Comments cũ
                    var elementsToClear = new FilteredElementCollector(doc)
                        .OfCategory(category)
                        .WhereElementIsNotElementType()
                        .ToList();
                    foreach (var element in elementsToClear)
                    {
                        element.LookupParameter("Mark")?.Set(string.Empty);
                        element.LookupParameter("Comments")?.Set(string.Empty);
                    }

                    // Gán giá trị Mark và Comments dựa trên ErrorMessages
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

                    // Filter: chỉ hiển thị các đối tượng có Mark != ""
                    if (markScheduleField != null)
                    {
                        ScheduleFilter filter = new ScheduleFilter(
                            schedule.Definition.GetField(0).FieldId,
                            ScheduleFilterType.NotEqual,
                            "");
                        schedule.Definition.AddFilter(filter);
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

                // Tạo/ lấy filter
                var parameterFilter = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .FirstOrDefault(f => f.Name == "DirectShape Filter");

                if (parameterFilter == null)
                    parameterFilter = ParameterFilterElement.Create(
                        doc,
                        "DirectShape Filter",
                        new List<ElementId> { new ElementId(BuiltInCategory.OST_GenericModel) });

                // Thêm filter vào View
                if (!view.IsFilterApplied(parameterFilter.Id))
                    view.AddFilter(parameterFilter.Id);

                // Tìm solid fill pattern
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

                // Cài đặt Override Graphic Settings
                OverrideGraphicSettings overrideSettings = new OverrideGraphicSettings();
                overrideSettings.SetSurfaceForegroundPatternColor(new Color(0, 255, 0)); // Màu xanh lá
                if (solidFillPatternId != null)
                    overrideSettings.SetSurfaceForegroundPatternId(solidFillPatternId);

                // Áp dụng cho các DirectShape
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
