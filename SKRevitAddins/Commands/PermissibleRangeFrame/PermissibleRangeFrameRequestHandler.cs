﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;
using SKRevitAddins.ViewModel;
using static SKRevitAddins.ViewModel.PermissibleRangeFrameViewModel;
using Document = Autodesk.Revit.DB.Document;
using UnitUtils = SKRevitAddins.Utils.UnitUtils;
using View = Autodesk.Revit.DB.View;

namespace SKRevitAddins.Commands.PermissibleRangeFrame
{
    public class PermissibleRangeFrameRequestHandler : IExternalEventHandler
    {
        private PermissibleRangeFrameViewModel ViewModel;
        public PermissibleRangeFrameRequestHandler(PermissibleRangeFrameViewModel viewModel)
        {
            ViewModel = viewModel;
        }
        private PermissibleRangeFrameRequest m_Request = new PermissibleRangeFrameRequest();

        public PermissibleRangeFrameRequest Request
        {
            get { return m_Request; }
        }

        public void Execute(UIApplication uiapp)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        break;
                    case (RequestId.OK):
                        CreatePermissibleRange(uiapp, ViewModel);
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("エラー", ex.Message);
            }
        }

        public string GetName()
        {
            return "";
        }

        #region Permissible Range
        public void CreatePermissibleRange(UIApplication uiapp, PermissibleRangeFrameViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var mepCurves = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfClass(typeof(MEPCurve))
                .Cast<MEPCurve>()
                .ToList();

            var x = viewModel.X;
            var y = viewModel.Y;
            var a = viewModel.A;
            var b = viewModel.B;
            var c = viewModel.C;

            var structuralFramings = viewModel.StructuralFramings;
            var sleeveSymbol = viewModel.SleeveSymbol;
            //var mepCurves = viewModel.MepCurves;
            var intersectionData = viewModel.IntersectionData;
            var errorMessages = viewModel.ErrorMessages;
            var sleevePlacements = viewModel.SleevePlacements;
            var directShapes = viewModel.DirectShapes;

            // Process intersections between frames and pipes/ducts and create direct shapes
            ProcessIntersections(doc, structuralFramings, mepCurves, intersectionData, directShapes, x, y);

            // Place sleeves at the intersection points
            PlaceSleeves(doc, sleeveSymbol, intersectionData, structuralFramings, sleevePlacements, errorMessages, directShapes, a, b, c);

            // If there are error messages, create error schedules
            if (viewModel.CreateErrorSchedules && errorMessages.Any())
            {
                CreateErrorSchedules(doc, errorMessages);
            }
            // Apply filters to the direct shapes
            ApplyFilterToDirectShapes(doc, uidoc.ActiveView, directShapes);


            using (Transaction subtx = new Transaction(doc, "ABC"))
            {
                subtx.Start();

                if (!viewModel.PlaceSleeves)
                {
                    // Collect all sleeves (FamilyInstances of category OST_PipeAccessory)
                    var sleeves = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_PipeAccessory)
                        .OfClass(typeof(FamilyInstance))
                        .WhereElementIsNotElementType()
                        .ToList();

                    // Collect their ElementIds for deletion
                    var sleeveIds = sleeves.Select(s => s.Id).ToList();

                    // Delete all sleeves
                    doc.Delete(sleeveIds);
                }

                if (!viewModel.PermissibleRange)
                {
                    // Collect the ElementIds of all direct shapes to delete
                    var directShapeIds = directShapes.Select(ds => ds.Id).ToList();

                    // Delete all direct shapes
                    doc.Delete(directShapeIds);
                }

                subtx.Commit();
            }

        }

        private void ProcessIntersections(Document doc, List<Element> structuralFramings, List<MEPCurve> mepCurves, Dictionary<(ElementId MEPCurveId, ElementId FrameId), List<XYZ>> intersectionData, List<DirectShape> directShapes, double x, double y)
        {
            using (Transaction trans = new Transaction(doc, "Tạo phạm vi cho phép xuyên dầm"))
            {
                trans.Start();

                foreach (var frame in structuralFramings)
                {
                    var frameObj = new FrameObj(frame);

                    // Get surrounding faces of the frame
                    var surroundingFaces = GetSurroundingFacesOfFrame(frameObj.FramingSolid);

                    // Create DirectShapes from the faces and add to directShapes list
                    foreach (Face face in surroundingFaces)
                    {
                        var directShape = CreateDirectShapeFromFrameFace(doc, frameObj, face, x, y);
                        if (directShape != null)
                        {
                            directShapes.Add(directShape);
                        }

                        foreach (var pipeOrDuct in mepCurves)
                        {
                            var pipeOrDuctCurve = (pipeOrDuct.Location as LocationCurve)?.Curve;
                            if (pipeOrDuctCurve == null) continue;

                            // Check for intersection between the face and the MEP curve
                            if (face.Intersect(pipeOrDuctCurve, out IntersectionResultArray resultArray) != SetComparisonResult.Overlap) continue;

                            var key = (MEPCurveId: pipeOrDuct.Id, FrameId: frame.Id);
                            if (!intersectionData.ContainsKey(key))
                            {
                                intersectionData[key] = new List<XYZ>();
                            }

                            // Add the intersection points to the intersection data
                            foreach (IntersectionResult intersectionResult in resultArray)
                            {
                                intersectionData[key].Add(intersectionResult.XYZPoint);
                            }
                        }
                    }
                }
                trans.Commit();
            }
        }

        private DirectShape CreateDirectShapeFromFrameFace(Document doc, FrameObj frameObj, Face face, double x, double y)
        {
            PlanarFace planarFace = face as PlanarFace;
            if (planarFace == null)
                return null;

            // Get the corners of the face
            BoundingBoxUV bboxUV = planarFace.GetBoundingBox();
            UV minUV = bboxUV.Min;
            UV maxUV = bboxUV.Max;

            // Calculate frame height and apply margins
            double frameHeight = frameObj.FramingHeight;
            double widthMargin = frameHeight * x;
            double heightMargin = frameHeight * y;

            UV adjustedMin = new UV(minUV.U + widthMargin, minUV.V + heightMargin);
            UV adjustedMax = new UV(maxUV.U - widthMargin, maxUV.V - heightMargin);

            // Ensure adjusted UVs are valid
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

            // Create profile for the face using new UV values, but only if the length is greater than Revit's tolerance
            List<Curve> profile = new List<Curve>();

            // Helper function to check curve length before creation
            Func<XYZ, XYZ, bool> isValidCurve = (start, end) => start.DistanceTo(end) > doc.Application.ShortCurveTolerance;

            XYZ corner1 = face.Evaluate(adjustedMin);
            XYZ corner2 = face.Evaluate(new UV(adjustedMin.U, adjustedMax.V));
            XYZ corner3 = face.Evaluate(adjustedMax);
            XYZ corner4 = face.Evaluate(new UV(adjustedMax.U, adjustedMin.V));

            // Check each edge length before adding it to the profile
            if (isValidCurve(corner1, corner2)) profile.Add(Line.CreateBound(corner1, corner2));
            if (isValidCurve(corner2, corner3)) profile.Add(Line.CreateBound(corner2, corner3));
            if (isValidCurve(corner3, corner4)) profile.Add(Line.CreateBound(corner3, corner4));
            if (isValidCurve(corner4, corner1)) profile.Add(Line.CreateBound(corner4, corner1));

            // If there are fewer than 3 valid curves, return null as the profile is invalid
            if (profile.Count < 3)
            {
                return null; // Not enough curves to create a valid loop
            }

            CurveLoop newCurveLoop = CurveLoop.Create(profile);

            // Extrusion direction typically normal to the face
            XYZ extrusionDirection = planarFace.FaceNormal;
            double extrusionDistance = 10.0 / 304.8;

            Solid directShapeSolid = GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { newCurveLoop }, extrusionDirection, extrusionDistance);

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
            //double minArea = faces.Min(f => f.Area);
            //var remainingFaces = faces
            //    .Where(f => f.Area > minArea) 
            //    .ToList();

            var faceAreas = faces.Select(face =>
            {
                Mesh mesh = face.Triangulate();
                double area = 0;
                int numTriangles = mesh.NumTriangles;
                for (int i = 0; i < numTriangles; i++)
                {
                    MeshTriangle triangle = mesh.get_Triangle(i);
                    XYZ p0 = triangle.get_Vertex(0);
                    XYZ p1 = triangle.get_Vertex(1);
                    XYZ p2 = triangle.get_Vertex(2);
                    area += 0.5 * ((p1 - p0).CrossProduct(p2 - p0)).GetLength();
                }
                return new { Face = face, Area = area };
            }).ToList();

            // Get the minimum area
            double minArea = faceAreas.Min(f => f.Area);

            // Filter out all faces that have the minimum area
            var remainingFaces = faceAreas
                .Where(f => f.Area > minArea)  // Remove all faces with the minimum area
                .Select(f => f.Face)
                .ToList();



            return remainingFaces;
        }

        private void PlaceSleeves(Document doc, FamilySymbol sleeveSymbol, Dictionary<(ElementId MEPCurveId, ElementId FrameId), List<XYZ>> intersectionData, List<Element> structuralFramings, Dictionary<ElementId, List<(XYZ, double)>> sleevePlacements, Dictionary<ElementId, HashSet<string>> errorMessages, List<DirectShape> directShapes, double a, double b, double c)
        {
            int successfulSleevesCount = 0; // Variable to count successfully placed sleeves

            // Collect all potential sleeves
            var potentialSleeves = new List<(ElementId MEPCurveId, XYZ Midpoint, double Diameter, double Length, XYZ Direction)>();

            // First pass: Collect all potential sleeves without creating them
            foreach (var entry in intersectionData)
            {
                var mepCurveId = entry.Key.MEPCurveId;
                var frameId = entry.Key.FrameId;
                var pipeOrDuct = doc.GetElement(mepCurveId);

                var frame = structuralFramings.FirstOrDefault(f => f.Id == frameId);
                if (frame == null) continue;

                var frameObj = new FrameObj(frame);
                var points = entry.Value;

                for (int i = 0; i + 1 < points.Count; i += 2)
                {
                    XYZ point1 = points[i];
                    XYZ point2 = points[i + 1];
                    XYZ midpoint = (point1 + point2) / 2;
                    XYZ direction = (point2 - point1).Normalize();

                    double pipeDiameter = pipeOrDuct.LookupParameter("Diameter")?.AsDouble() ?? 0;
                    double sleeveDiameter = pipeDiameter + UnitUtils.MmToFeet(50);
                    double frameHeight = frameObj.FramingHeight;

                    // Perform initial checks
                    if (sleeveDiameter > UnitUtils.MmToFeet(a))
                    {
                        if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                        {
                            errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                        }
                        errorMessages[pipeOrDuct.Id].Add("OD > 750mm");
                        continue;
                    }

                    if (sleeveDiameter > frameHeight * b)
                    {
                        if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                        {
                            errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                        }
                        errorMessages[pipeOrDuct.Id].Add("OD > H/3");
                        continue;
                    }

                    double sleeveLength = point1.DistanceTo(point2);

                    // Add the potential sleeve to the list
                    potentialSleeves.Add((MEPCurveId: mepCurveId, Midpoint: midpoint, Diameter: sleeveDiameter, Length: sleeveLength, Direction: direction));
                }
            }

            // Second pass: Check for sleeves that are too close and remove both from the list
            var sleevesToRemove = new HashSet<int>();
            for (int i = 0; i < potentialSleeves.Count; i++)
            {
                for (int j = i + 1; j < potentialSleeves.Count; j++)
                {
                    var sleeve1 = potentialSleeves[i];
                    var sleeve2 = potentialSleeves[j];

                    // Calculate the minimum allowed distance between the sleeves
                    double minDistance = (sleeve1.Diameter + sleeve2.Diameter) * (c);

                    // Calculate the actual distance between the current and previous midpoint
                    double actualDistance = sleeve1.Midpoint.DistanceTo(sleeve2.Midpoint);

                    if (!(actualDistance < minDistance)) continue;
                    // Mark both sleeves for removal
                    sleevesToRemove.Add(i);
                    sleevesToRemove.Add(j);

                    // Record error messages for both sleeves
                    if (!errorMessages.ContainsKey(sleeve1.MEPCurveId))
                    {
                        errorMessages[sleeve1.MEPCurveId] = new HashSet<string>();
                    }
                    errorMessages[sleeve1.MEPCurveId].Add("Khoảng cách giữa hai Sleeve < (OD1 + OD2)*2/3");

                    if (!errorMessages.ContainsKey(sleeve2.MEPCurveId))
                    {
                        errorMessages[sleeve2.MEPCurveId] = new HashSet<string>();
                    }
                    errorMessages[sleeve2.MEPCurveId].Add("Khoảng cách giữa hai Sleeve < (OD1 + OD2)*2/3");
                }
            }

            // Remove sleeves that are too close
            potentialSleeves = potentialSleeves.Where((sleeve, index) => !sleevesToRemove.Contains(index)).ToList();

            // Third pass: Create the sleeves that passed all checks
            using (Transaction trans = new Transaction(doc, "Đặt Sleeve"))
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

                    // Create the sleeve instance

                    FamilyInstance sleeveInstance = doc.Create.NewFamilyInstance(midpoint, sleeveSymbol, StructuralType.NonStructural);
                    // Rotate the sleeve to align with the pipe/duct direction
                    Line rotationAxis = Line.CreateBound(midpoint, midpoint + XYZ.BasisZ);
                    double rotationAngle = XYZ.BasisX.AngleTo(direction) + Math.PI / 2;
                    ElementTransformUtils.RotateElement(doc, sleeveInstance.Id, rotationAxis, rotationAngle);

                    // Set the sleeve parameters
                    Parameter lengthParam = sleeveInstance.LookupParameter("L");
                    if (lengthParam != null)
                    {
                        lengthParam.Set(sleeveLength);
                    }

                    Parameter odParam = sleeveInstance.LookupParameter("OD");
                    if (odParam != null)
                    {
                        odParam.Set(sleeveDiameter);
                    }

                    // Record the sleeve placement
                    if (!sleevePlacements.ContainsKey(mepCurveId))
                    {
                        sleevePlacements[mepCurveId] = new List<(XYZ, double)>();
                    }
                    sleevePlacements[mepCurveId].Add((midpoint, sleeveDiameter));

                    // Increase the count for successfully placed sleeves
                    successfulSleevesCount++;

                    // Check if the sleeve is within the permissible range
                    bool isWithinDirectShapes = directShapes.All(ds => IsPointWithinDirectShape(midpoint, ds));
                    if (!isWithinDirectShapes) continue;
                    if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                    {
                        errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                    }
                    errorMessages[pipeOrDuct.Id].Add("Sleeve nằm ngoài phạm vi cho phép xuyên dầm.");
                }

                trans.Commit();
            }

            // Update TaskDialog to show the number of successful sleeves
            TaskDialog.Show("Thông báo:", $"Đã đặt {successfulSleevesCount} Sleeves thỏa điều kiện. Vui lòng kiểm tra lại.");
        }

        private bool IsPointWithinDirectShape(XYZ point, DirectShape directShape)
        {
            // Obtain the geometry of the DirectShape
            GeometryElement geometryElement = directShape.get_Geometry(new Options());

            // Iterate through the solids in the geometry
            foreach (GeometryObject geomObj in geometryElement)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                {
                    // Check if the point is contained within the solid by projecting onto the solid faces
                    var faceArray = solid.Faces;
                    foreach (Face face in faceArray)
                    {
                        var result = face.Project(point);
                        if (result != null && result.Distance < 0.001)
                        {
                            return true; // Point is within the solid
                        }
                    }
                }
            }
            return false; // Point is not within the solid
        }
        private void CreateErrorSchedules(Document doc, Dictionary<ElementId, HashSet<string>> errorMessages)
        {
            using (Transaction subtx = new Transaction(doc, "Tạo Report"))
            {
                subtx.Start();

                // List of schedule names and categories for pipes and ducts
                var schedules = new List<(string scheduleName, BuiltInCategory category)>
                {
                    ("PipeErrorSchedule", BuiltInCategory.OST_PipeCurves),
                    ("DuctErrorSchedule", BuiltInCategory.OST_DuctCurves)
                };

                foreach (var (scheduleName, category) in schedules)
                {
                    // Check and delete existing schedule with the same name
                    var existingSchedule = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSchedule))
                        .Cast<ViewSchedule>()
                        .FirstOrDefault(sch => sch.Name.Equals(scheduleName));

                    if (existingSchedule != null)
                    {
                        doc.Delete(existingSchedule.Id); // Delete the existing schedule
                    }

                    // Create a new schedule
                    ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, new ElementId(category));
                    schedule.Name = scheduleName;

                    // Get the schedulable fields
                    var schedulableFields = schedule.Definition.GetSchedulableFields();

                    // Add "Mark" and "Comments" fields to the schedule
                    SchedulableField markField = schedulableFields.FirstOrDefault(sf => sf.GetName(doc) == "Mark");
                    SchedulableField commentField = schedulableFields.FirstOrDefault(sf => sf.GetName(doc) == "Comments");

                    // Add fields and store the ScheduleField objects
                    ScheduleField markScheduleField = null;
                    if (markField != null)
                    {
                        markScheduleField = schedule.Definition.AddField(markField);
                    }

                    ScheduleField commentScheduleField = null;
                    if (commentField != null)
                    {
                        commentScheduleField = schedule.Definition.AddField(commentField);
                    }

                    // Clear old Mark and Comments values
                    var elementsToClear = new FilteredElementCollector(doc)
                        .OfCategory(category)
                        .WhereElementIsNotElementType()
                        .ToList();

                    foreach (var element in elementsToClear)
                    {
                        Parameter markParam = element.LookupParameter("Mark");
                        if (markParam != null)
                        {
                            markParam.Set(string.Empty);
                        }

                        Parameter commentParam = element.LookupParameter("Comments");
                        if (commentParam != null)
                        {
                            commentParam.Set(string.Empty);
                        }
                    }

                    // Set new values for Mark and Comments based on errorMessages
                    int markIndex = 1;
                    foreach (var kvp in errorMessages)
                    {
                        Element element = doc.GetElement(kvp.Key);
                        if (element.Category.Id.IntegerValue != (int)category)
                            continue;

                        Parameter markParam = element.LookupParameter("Mark");
                        if (markParam != null && kvp.Value.Any())
                        {
                            markParam.Set(markIndex.ToString());
                            markIndex++;
                        }

                        Parameter commentParam = element.LookupParameter("Comments");
                        if (commentParam != null)
                        {
                            commentParam.Set(string.Join(", ", kvp.Value));
                        }
                    }

                    // Apply filter to remove elements without errors
                    if (markScheduleField == null) continue;

                    // Use ScheduleFilter.Create method for string comparison
                    ScheduleFilter filter = new ScheduleFilter(schedule.Definition.GetField(0).FieldId, ScheduleFilterType.NotEqual, "");
                    schedule.Definition.AddFilter(filter);
                }

                subtx.Commit();
            }
        }

        private void ApplyFilterToDirectShapes(Document doc, View view, List<DirectShape> directShapes)
        {
            using (Transaction trans = new Transaction(doc, "Tạo filter thể hiện phạm vi xuyên dầm"))
            {
                trans.Start();

                // Get or create the parameter filter
                ParameterFilterElement parameterFilter = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .FirstOrDefault(f => f.Name == "DirectShape Filter");

                if (parameterFilter == null)
                {
                    parameterFilter = ParameterFilterElement.Create(doc, "DirectShape Filter", new List<ElementId> { new ElementId(BuiltInCategory.OST_GenericModel) });
                }

                if (!view.IsFilterApplied(parameterFilter.Id))
                {
                    view.AddFilter(parameterFilter.Id);
                }

                // Get solid fill pattern
                ElementId solidFillPatternId = null;
                var fillPatternElements = new FilteredElementCollector(doc)
                    .OfClass(typeof(FillPatternElement))
                    .Cast<FillPatternElement>();

                foreach (var fp in fillPatternElements)
                {
                    if (!fp.GetFillPattern().IsSolidFill) continue;
                    solidFillPatternId = fp.Id;
                    break;
                }

                // Set override graphic settings
                OverrideGraphicSettings overrideSettings = new OverrideGraphicSettings();
                overrideSettings.SetSurfaceForegroundPatternColor(new Color(0, 255, 0)); // Green color
                if (solidFillPatternId != null)
                {
                    overrideSettings.SetSurfaceForegroundPatternId(solidFillPatternId);
                }

                // Apply overrides to the direct shapes
                foreach (var directShape in directShapes)
                {
                    view.SetElementOverrides(directShape.Id, overrideSettings);
                }
                trans.Commit();
            }
        }

        #endregion
    }
}
