﻿using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;
using Curve = Autodesk.Revit.DB.Curve;
using Face = Autodesk.Revit.DB.Face;
using Line = Autodesk.Revit.DB.Line;
using Solid = Autodesk.Revit.DB.Solid;
using Transaction = Autodesk.Revit.DB.Transaction;
using UnitUtils = Autodesk.Revit.DB.UnitUtils;

namespace SKRevitAddins.Commands.PermissibleRangeFramePunching
{
    [Transaction(TransactionMode.Manual)]
    public class PermissibleRangeFramePunchingCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = uidoc.ActiveView;

            var linkedDocs = GetLinkedDocuments(doc);
            var pipesAndDucts = GetElementsOfType<MEPCurve>(doc).Where(e => e.Document.ActiveView.Id == activeView.Id).ToList();
            var structuralFramings = new List<Element>();

            var intersectionData = new Dictionary<ElementId, List<XYZ>>();
            var errorMessages = new Dictionary<ElementId, HashSet<string>>();
            var sleevePlacements = new Dictionary<ElementId, List<(XYZ, double)>>();
            var directShapes = new List<DirectShape>();
            List<FrameObj> frameObjs = new List<FrameObj>();
            List<Solid> expandedFramingSolids = new List<Solid>();
            Solid unionExpandedFramingSolids = null;
            foreach (var linkedDoc in linkedDocs)
            {
                structuralFramings.AddRange(GetElementsOfType<FamilyInstance>(linkedDoc, BuiltInCategory.OST_StructuralFraming));
            }

            foreach (var framing in structuralFramings)
            {
                FrameObj frameObj = new FrameObj(framing);
                frameObjs.Add(frameObj);
                ProcessIntersections(doc, frameObj, pipesAndDucts, intersectionData, directShapes);
                //expandedFramingSolids.Add(frameObj.FramingSolidExpand);
                //unionExpandedFramingSolids = expandedFramingSolids.UnionSolidList();
            }

            var sleeveSymbol = GetSleeveSymbol(doc, ref message);
            if (sleeveSymbol == null)
                return Result.Failed;

            PlaceSleeves(doc, sleeveSymbol, intersectionData, structuralFramings, sleevePlacements, errorMessages, directShapes);

            if (errorMessages.Any()) CreateErrorSchedules(doc, errorMessages);

            TaskDialog.Show("Intersections", $"Placed {intersectionData.Count} スリーブ_SK instances at intersections.");

            ApplyFilterToDirectShapes(doc, activeView, directShapes);

            return Result.Succeeded;
        }

        private List<Document> GetLinkedDocuments(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Select(link => link.GetLinkDocument())
                .Where(linkedDoc => linkedDoc != null)
                .ToList();
        }

        private List<Element> GetElementsOfType<T>(Document doc, BuiltInCategory? category = null) where T : Element
        {
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(T))
                .WhereElementIsNotElementType();

            if (category.HasValue)
                collector.OfCategory(category.Value);

            return collector.ToElements().ToList();
        }

        // Process intersections between structural framings and pipes/ducts
        private void ProcessIntersections(Document doc, FrameObj frame, List<Element> pipesAndDucts, Dictionary<ElementId, List<XYZ>> intersectionData, List<DirectShape> directShapes)
        {
            using (Transaction trans = new Transaction(doc, "Place Sleeves and Create Direct Shapes"))
            {
                trans.Start();
                //mergedExpandedFramingSolid.BakeSolidToDirectShape(doc);
                var surroundingFaces = GetSurroundingFaces(frame.FramingSolid);
                foreach (Face face in surroundingFaces)
                {
                    //var directShape = CreateDirectShapeFromFrameFace(doc, frame, face, mergedExpandedFramingSolid);
                    var directShape = CreateDirectShapeFromFrameFace(doc, frame, face);
                    if (directShape != null)
                    {
                        directShapes.Add(directShape);
                    }
                    foreach (var pipeOrDuct in pipesAndDucts)
                    {
                        var pipeOrDuctCurve = (pipeOrDuct.Location as LocationCurve)?.Curve;
                        if (pipeOrDuctCurve == null)
                            continue;

                        if (face.Intersect(pipeOrDuctCurve, out IntersectionResultArray resultArray) == SetComparisonResult.Overlap)
                        {
                            if (!intersectionData.ContainsKey(pipeOrDuct.Id))
                            {
                                intersectionData[pipeOrDuct.Id] = new List<XYZ>();
                            }

                            foreach (IntersectionResult intersectionResult in resultArray)
                            {
                                intersectionData[pipeOrDuct.Id].Add(intersectionResult.XYZPoint);
                            }
                        }
                    }
                }
                // Tiếp tục quá trình tạo DirectShape với mergedExpandedFramingSolid...

                trans.Commit();
            }
        }

        // Retrieve sleeve family symbol
        private FamilySymbol GetSleeveSymbol(Document doc, ref string message)
        {
            var sleeveSymbol = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeAccessory)
                .OfClass(typeof(FamilySymbol))
                .WhereElementIsElementType()
                .Cast<FamilySymbol>()
                .FirstOrDefault(symbol => symbol.FamilyName == "スリーブ_SK");

            if (sleeveSymbol == null)
            {
                message = "The スリーブ_SK family could not be found.";
                return null;
            }

            if (!sleeveSymbol.IsActive)
            {
                sleeveSymbol.Activate();
                doc.Regenerate();
            }

            return sleeveSymbol;
        }

        // Place sleeves at intersection points
        private void PlaceSleeves(Document doc, FamilySymbol sleeveSymbol, Dictionary<ElementId, List<XYZ>> intersectionData, List<Element> structuralFramings, Dictionary<ElementId, List<(XYZ, double)>> sleevePlacements, Dictionary<ElementId, HashSet<string>> errorMessages, List<DirectShape> directShapes)
        {
            using (Transaction trans = new Transaction(doc, "Place Sleeves"))
            {
                trans.Start();

                foreach (var entry in intersectionData)
                {
                    var pipeOrDuct = doc.GetElement(entry.Key);
                    var points = entry.Value;

                    for (int i = 0; i < points.Count; i += 2)
                    {
                        if (i + 1 < points.Count)
                        {
                            XYZ point1 = points[i];
                            XYZ point2 = points[i + 1];
                            XYZ midpoint = (point1 + point2) / 2;
                            XYZ direction = (point2 - point1).Normalize();

                            double pipeDiameter = pipeOrDuct.LookupParameter("Diameter")?.AsDouble() ?? 0;
                            double sleeveDiameter = pipeDiameter + UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Millimeters);
                            double beamHeight = GetBeamHeight(point1, structuralFramings);

                            HashSet<string> errors = ValidateSleevePlacement(sleeveDiameter, beamHeight, midpoint, sleevePlacements, entry.Key);

                            if (errors.Count > 0)
                            {
                                if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                                {
                                    errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                                }
                                errorMessages[pipeOrDuct.Id].UnionWith(errors);
                                continue;
                            }

                            // Kiểm tra nếu pipe/duct không nằm hoàn toàn trong phạm vi direct shape sẽ không được tạo ra
                            if (!IsPointWithinAnyDirectShape(point1, directShapes) || !IsPointWithinAnyDirectShape(point2, directShapes))
                            {
                                if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                                {
                                    errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                                }
                                errorMessages[pipeOrDuct.Id].Add("Pipe/Duct does not fully lie within the permissible direct shape boundaries.");
                                continue;
                            }

                            PlaceSleeveInstance(doc, sleeveSymbol, midpoint, direction, point1, point2, sleeveDiameter, sleevePlacements, entry.Key, pipeOrDuct, directShapes, errorMessages);
                        }
                    }
                }

                trans.Commit();
            }
        }

        // Validate sleeve placement conditions
        private HashSet<string> ValidateSleevePlacement(double sleeveDiameter, double beamHeight, XYZ midpoint, Dictionary<ElementId, List<(XYZ, double)>> sleevePlacements, ElementId pipeOrDuctId)
        {
            HashSet<string> errors = new HashSet<string>();

            if (sleeveDiameter > UnitUtils.ConvertToInternalUnits(750, UnitTypeId.Millimeters))
            {
                errors.Add("OD > 750mm");
            }

            if (sleeveDiameter > beamHeight / 3)
            {
                errors.Add("OD > H/3");
            }

            if (!sleevePlacements.ContainsKey(pipeOrDuctId))
            {
                sleevePlacements[pipeOrDuctId] = new List<(XYZ, double)>();
            }

            foreach (var (otherMidpoint, otherDiameter) in sleevePlacements[pipeOrDuctId])
            {
                double minDistance = (sleeveDiameter + otherDiameter) * 1.5;
                if (Math.Abs(midpoint.X - otherMidpoint.X) < minDistance || Math.Abs(midpoint.Y - otherMidpoint.Y) < minDistance)
                {
                    errors.Add($"Distance between sleeves < (OD1 + OD2)*3/2");
                    break;
                }
            }

            return errors;
        }

        // Place sleeve instance at midpoint
        private void PlaceSleeveInstance(Document doc, FamilySymbol sleeveSymbol, XYZ midpoint, XYZ direction, XYZ point1, XYZ point2, double sleeveDiameter, Dictionary<ElementId, List<(XYZ, double)>> sleevePlacements, ElementId entryKey, Element pipeOrDuct, List<DirectShape> directShapes, Dictionary<ElementId, HashSet<string>> errorMessages)
        {
            FamilyInstance sleeveInstance = doc.Create.NewFamilyInstance(midpoint, sleeveSymbol, StructuralType.NonStructural);

            Line axis = Line.CreateBound(midpoint, midpoint + XYZ.BasisZ);
            double angle = XYZ.BasisX.AngleTo(direction);
            double additionalRotation = Math.PI / 2;
            ElementTransformUtils.RotateElement(doc, sleeveInstance.Id, axis, angle + additionalRotation);

            Parameter lengthParam = sleeveInstance.LookupParameter("L");
            if (lengthParam != null)
            {
                lengthParam.Set(point1.DistanceTo(point2));
            }

            Parameter odParam = sleeveInstance.LookupParameter("OD");
            if (odParam != null)
            {
                odParam.Set(sleeveDiameter);
            }

            sleevePlacements[entryKey].Add((midpoint, sleeveDiameter));

            bool isWithinDirectShape = directShapes.Any(ds => IsPointWithinDirectShape(midpoint, ds));
            if (!isWithinDirectShape)
            {
                if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                {
                    errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                }
                errorMessages[pipeOrDuct.Id].Add("Sleeve Placed Outside Permissible Beam Penetration Range.");
            }
        }

        // Handle errors and prompt to save error messages
        private void CreateErrorSchedules(Document doc, Dictionary<ElementId, HashSet<string>> errorMessages)
        {
            // Create a schedule for pipes
            CreateErrorSchedule(doc, errorMessages, "PipeErrorSchedule", BuiltInCategory.OST_PipeCurves);

            // Create a schedule for ducts
            CreateErrorSchedule(doc, errorMessages, "DuctErrorSchedule", BuiltInCategory.OST_DuctCurves);
        }

        private void CreateErrorSchedule(Document doc, Dictionary<ElementId, HashSet<string>> errorMessages, string scheduleName, BuiltInCategory category)
        {
            using (Transaction trans = new Transaction(doc, "Create Error Schedule"))
            {
                trans.Start();

                ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, new ElementId(category));

                schedule.Name = scheduleName;

                SchedulableField markField = schedule.Definition.GetSchedulableFields().First(sf => sf.GetName(doc) == "Mark");
                SchedulableField commentField = schedule.Definition.GetSchedulableFields().First(sf => sf.GetName(doc) == "Comments");

                schedule.Definition.AddField(ScheduleFieldType.Instance, markField.ParameterId);
                schedule.Definition.AddField(ScheduleFieldType.Instance, commentField.ParameterId);

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
                ScheduleFilter filter = new ScheduleFilter(schedule.Definition.GetField(0).FieldId, ScheduleFilterType.Equal, "");
                schedule.Definition.AddFilter(filter);

                trans.Commit();
            }
        }

        // Get beam height from structural framings
        private double GetBeamHeight(XYZ point, List<Element> structuralFramings)
        {
            foreach (var framing in structuralFramings)
            {
                if (IsPointOnElement(framing, point))
                {
                    BoundingBoxXYZ boundingBox = framing.get_BoundingBox(null);
                    if (boundingBox != null)
                    {
                        return boundingBox.Max.Z - boundingBox.Min.Z;
                    }
                }
            }
            return 0;
        }

        // Check if a point is on a given element
        private bool IsPointOnElement(Element element, XYZ point)
        {
            var boundingBox = element.get_BoundingBox(null);
            return boundingBox != null &&
                   point.X >= boundingBox.Min.X && point.X <= boundingBox.Max.X &&
                   point.Y >= boundingBox.Min.Y && point.Y <= boundingBox.Max.Y &&
                   point.Z >= boundingBox.Min.Z && point.Z <= boundingBox.Max.Z;
        }

        // Get solids from geometry element
        private static List<Solid> GetSolidsFromGeometry(GeometryElement geometryElement)
        {
            List<Solid> solids = new List<Solid>();

            foreach (GeometryObject geomObj in geometryElement)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                {
                    solids.Add(solid);
                }
                else if (geomObj is GeometryInstance geomInstance)
                {
                    GeometryElement instanceGeometry = geomInstance.GetInstanceGeometry();
                    solids.AddRange(GetSolidsFromGeometry(instanceGeometry));
                }
            }

            return solids;
        }

        private static List<Face> GetSurroundingFaces(Solid solid)
        {
            var faces = solid.GetSolidVerticalFaces();

            var faceAreas = faces.Select(face => new { Face = face, Area = GetFaceArea(face) }).ToList();

            var sortedFaceAreas = faceAreas.OrderBy(f => f.Area).ToList();

            double minArea = sortedFaceAreas.First().Area;

            var remainingFaces = sortedFaceAreas
                .Where(f => f.Area > minArea)  
                .Select(f => f.Face)           
                .ToList();

            return remainingFaces;
        }

        // Get area of a face
        private static double GetFaceArea(Face face)
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

            return area;
        }

        //private DirectShape CreateDirectShapeFromFrameFace(Document doc, FrameObj frameObj, Face face)
        //{
        //    // Get BoundingBoxUV 
        //    BoundingBoxUV boundingBox = face.GetBoundingBox();
        //    UV min = boundingBox.Min;
        //    UV max = boundingBox.Max;

        //    // get BoundingBoxXYZ of solid to calc frameHeight
        //    BoundingBoxXYZ solidBoundingBox = frameObj.FramingSolid.GetBoundingBox();

        //    double heightMargin = frameObj.FramingHeight / 4;
        //    double widthMargin = frameObj.FramingHeight / 2;

        //    // Xác định biên
        //    UV adjustedMin = new UV(min.U + widthMargin, min.V + heightMargin);
        //    UV adjustedMax = new UV(max.U - widthMargin, max.V - heightMargin);

        //    // Nếu điều chỉnh vượt ra ngoài biên, thì đặt lại giá trị hợp lý cho U, V
        //    if (adjustedMin.V >= adjustedMax.V)
        //    {
        //        adjustedMin = new UV(adjustedMin.U, min.V + (max.V - min.V) / 4);
        //        adjustedMax = new UV(adjustedMax.U, max.V - (max.V - min.V) / 4);
        //    }
        //    if (adjustedMin.U >= adjustedMax.U)
        //    {
        //        adjustedMin = new UV(min.U + (max.U - min.U) / 4, adjustedMin.V);
        //        adjustedMax = new UV(max.U - (max.U - min.U) / 4, adjustedMax.V);
        //    }

        //    // Tạo profile cho face bằng cách sử dụng các giá trị UV mới
        //    List<Curve> profile = new List<Curve>
        //    {
        //        Line.CreateBound(face.Evaluate(adjustedMin), face.Evaluate(new UV(adjustedMin.U, adjustedMax.V))),
        //        Line.CreateBound(face.Evaluate(new UV(adjustedMin.U, adjustedMax.V)), face.Evaluate(adjustedMax)),
        //        Line.CreateBound(face.Evaluate(adjustedMax), face.Evaluate(new UV(adjustedMax.U, adjustedMin.V))),
        //        Line.CreateBound(face.Evaluate(new UV(adjustedMax.U, adjustedMin.V)), face.Evaluate(adjustedMin))
        //    };

        //    // Tạo CurveLoop từ profile
        //    CurveLoop curveLoop = CurveLoop.Create(profile);
        //    List<CurveLoop> curveLoops = new List<CurveLoop> { curveLoop };
        //    XYZ extrusionDirection = face.ComputeNormal(UV.Zero);

        //    // Tạo khối directShapeSolid ban đầu
        //    Solid directShapeSolid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoops, extrusionDirection, 10.0 / 304.8);

        //    //// Kiểm tra nếu mergedExpandedFramingSolid có giao cắt với directShapeSolid trước khi thực hiện Boolean Difference
        //    //if (mergeSolid != null && directShapeSolid != null && mergeSolid.Volume > 0 && directShapeSolid.Volume > 0)
        //    //{
        //    //    try
        //    //    {
        //    //        directShapeSolid = BooleanOperationsUtils.ExecuteBooleanOperation(directShapeSolid, mergeSolid, BooleanOperationsType.Difference);
        //    //    }
        //    //    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
        //    //    {
        //    //        // Xử lý nếu phép toán boolean không hợp lệ
        //    //    }
        //    //}

        //    if (directShapeSolid != null && directShapeSolid.Volume > 0)
        //    {
        //        DirectShape directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
        //        directShape.SetShape(new GeometryObject[] { directShapeSolid });
        //        directShape.SetName("Beam Intersection Zone");

        //        return directShape;
        //    }

        //    // Trả về null nếu khối đã bị loại bỏ hoàn toàn bởi giao cắt
        //    return null;
        //}
        private DirectShape CreateDirectShapeFromFrameFace(Document doc, FrameObj frameObj, Face face)
        {
            // Ensure the face is a PlanarFace before proceeding
            PlanarFace planarFace = face as PlanarFace;
            if (planarFace == null)
                return null;

            // Get the face's edges as curve loops
            IList<CurveLoop> faceEdges = planarFace.GetEdgesAsCurveLoops();
            if (faceEdges == null || faceEdges.Count == 0)
                return null;

            double frameHeight = frameObj.FramingHeight;
            double heightMargin = frameHeight / 4;
            double widthMargin = frameHeight / 2;

            // Apply offset to the first curve loop to create a smaller profile
            CurveLoop offsetLoop = CurveLoop.CreateViaOffset(faceEdges.First(), - widthMargin, planarFace.FaceNormal);
            List<CurveLoop> curveLoops = new List<CurveLoop> { offsetLoop };
            XYZ extrusionDirection = planarFace.FaceNormal;

            // Create the direct shape solid using the offset loop and the frame height
            Solid directShapeSolid = null;
            try
            {
                // Extrude the curve loop along the face normal direction with a fixed height
                directShapeSolid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoops, extrusionDirection, 10.0 / 304.8);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException ex)
            {
                TaskDialog.Show("Error", $"Extrusion failed: {ex.Message}");
                return null;
            }

            // Check if the resulting solid is valid
            if (directShapeSolid != null && directShapeSolid.Volume > 0)
            {
                DirectShape directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                directShape.SetShape(new GeometryObject[] { directShapeSolid });
                directShape.SetName("Beam Intersection Zone");

                return directShape;
            }

            // Return null if the solid has been completely removed or is invalid
            return null;
        }

        // Manually validate the CurveLoop
        private bool IsCurveLoopValid(CurveLoop curveLoop)
        {
            if (curveLoop == null || curveLoop.Count() == 0)
                return false;

            XYZ lastEndPoint = null;

            foreach (Curve curve in curveLoop)
            {
                if (curve == null)
                    return false;

                // Get the start and end points of the current curve
                XYZ startPoint = curve.GetEndPoint(0);
                XYZ endPoint = curve.GetEndPoint(1);

                // Check if the curve's start point matches the end point of the previous curve (if any)
                if (lastEndPoint != null && !startPoint.IsAlmostEqualTo(lastEndPoint))
                {
                    return false; // The loop is not continuous
                }

                // Set the end point of the current curve as the last end point
                lastEndPoint = endPoint;
            }

            // Finally, check if the last end point matches the start point of the first curve (the loop must be closed)
            return lastEndPoint.IsAlmostEqualTo(curveLoop.First().GetEndPoint(0));
        }

        // Check if a point is within a direct shape
        private bool IsPointWithinDirectShape(XYZ point, DirectShape directShape)
        {
            var boundingBox = directShape.get_BoundingBox(null);
            if (boundingBox == null)
                return false;

            return (point.X >= boundingBox.Min.X && point.X <= boundingBox.Max.X) &&
                   (point.Y >= boundingBox.Min.Y && point.Y <= boundingBox.Max.Y) &&
                   (point.Z >= boundingBox.Min.Z && point.Z <= boundingBox.Max.Z);
        }

        // Check if a point is within any of the direct shapes
        private bool IsPointWithinAnyDirectShape(XYZ point, List<DirectShape> directShapes)
        {
            return directShapes.Any(ds => IsPointWithinDirectShape(point, ds));
        }

        // Apply filter to direct shapes and set color
        private void ApplyFilterToDirectShapes(Document doc, View view, List<DirectShape> directShapes)
        {
            using (Transaction trans = new Transaction(doc, "Apply Filter to Direct Shapes"))
            {
                trans.Start();

                ParameterFilterElement parameterFilter = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .FirstOrDefault(f => f.Name == "DirectShape Filter");

                if (parameterFilter == null)
                {
                    ElementCategoryFilter directShapeFilter = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel);
                    parameterFilter = ParameterFilterElement.Create(doc, "DirectShape Filter", new List<ElementId> { new ElementId(BuiltInCategory.OST_GenericModel) });
                    view.AddFilter(parameterFilter.Id); 
                }
                else if (!view.IsFilterApplied(parameterFilter.Id)) 
                {
                    view.AddFilter(parameterFilter.Id); 
                }

                ElementId solidFillPatternId = null;
                List<FillPatternElement> fillPatternList = new FilteredElementCollector(doc)
                    .WherePasses(new ElementClassFilter(typeof(FillPatternElement))).
                    ToElements().Cast<FillPatternElement>().ToList();

                foreach (FillPatternElement fp in fillPatternList)
                {
                    if (fp.GetFillPattern().IsSolidFill)
                    {
                        solidFillPatternId = fp.Id;
                        break;
                    }
                }

                OverrideGraphicSettings overrideSettings = new OverrideGraphicSettings();
                overrideSettings.SetSurfaceForegroundPatternColor(new Color(0, 255, 0));
                overrideSettings.SetSurfaceForegroundPatternId(solidFillPatternId);

                foreach (var directShape in directShapes)
                {
                    view.SetElementOverrides(directShape.Id, overrideSettings);
                }

                trans.Commit();
            }
        }
        public class FrameObj
        {
            public Element FramingObj { get; set; }
            public GeometryElement FramingGeometryObject { get; set; }
            public Solid FramingSolid { get; set; }
            public double FramingHeight { get; set; }
            private List<Face> framingFaces { get; set; }

            public FrameObj(Element frameObj)
            {
                FramingObj = frameObj;
                FramingGeometryObject = frameObj.get_Geometry(new Options());
                if (FramingGeometryObject == null) return;
                List<Solid> solids = GetSolidsFromGeometry(FramingGeometryObject);
                FramingSolid = solids.UnionSolidList();
                FramingHeight = FramingSolid.GetSolidHeight();
                framingFaces = GetSurroundingFaces(FramingSolid);
            }

            //private Solid ExpandSolidByOffset(Solid originalSolid, double offsetAmount)
            //{
            //    if (originalSolid == null || originalSolid.Volume == 0)
            //        return null;

            //    // Initialize a list to hold the new solids created from the offset faces
            //    List<Solid> expandedSolids = new List<Solid>();

            //    foreach (Face face in originalSolid.Faces)
            //    {
            //        PlanarFace planarFace = face as PlanarFace;
            //        if (planarFace == null) continue;

            //        IList<CurveLoop> faceEdges = planarFace.GetEdgesAsCurveLoops();
            //        if (faceEdges == null || faceEdges.Count == 0) continue;
            //        CurveLoop offsetLoop = CurveLoop.CreateViaOffset(faceEdges.First(), offsetAmount, planarFace.FaceNormal);
            //        List<CurveLoop> curveLoops = new List<CurveLoop> { offsetLoop };
            //        Solid extrudedSolid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoops, planarFace.FaceNormal, offsetAmount);

            //        if (extrudedSolid != null)
            //        {
            //            expandedSolids.Add(extrudedSolid);
            //        }
            //    }
            //    return expandedSolids.UnionSolidList();
            //}
        }

    }
}
