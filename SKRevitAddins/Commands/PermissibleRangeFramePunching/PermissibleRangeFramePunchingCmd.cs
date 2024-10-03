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
            using (Transaction trans = new Transaction(doc, "Place Sleeves and Create Direct Shapes"))
            {
                trans.Start();
                foreach (var framing in structuralFramings)
                {
                    FrameObj frameObj = new FrameObj(framing);
                    frameObjs.Add(frameObj);
                    ProcessIntersections(doc, frameObj, pipesAndDucts, intersectionData, directShapes);
                }

                var sleeveSymbol = GetSleeveSymbol(doc, ref message);
                if (sleeveSymbol == null)
                    return Result.Failed;

                PlaceSleeves(doc, sleeveSymbol, intersectionData, structuralFramings, sleevePlacements, errorMessages, directShapes);

                if (errorMessages.Any()) CreateErrorSchedules(doc, errorMessages);

                TaskDialog.Show("Intersections", $"Placed {intersectionData.Count} スリーブ_SK instances at intersections.");

                ApplyFilterToDirectShapes(doc, activeView, directShapes);
                trans.Commit();
            }
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
            var surroundingFaces = GetSurroundingFaces(frame.FramingSolid);
            foreach (Face face in surroundingFaces)
            {
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

            if (sleeveSymbol != null)
            {
                if (sleeveSymbol.IsActive) return sleeveSymbol;
                sleeveSymbol.Activate();
                doc.Regenerate();

                return sleeveSymbol;
            }

            message = "The スリーブ_SK family could not be found.";
            return null;
        }

        // Place sleeves at intersection points
        private void PlaceSleeves(Document doc, FamilySymbol sleeveSymbol, Dictionary<ElementId, List<XYZ>> intersectionData, List<Element> structuralFramings, Dictionary<ElementId, List<(XYZ, double)>> sleevePlacements, Dictionary<ElementId, HashSet<string>> errorMessages, List<DirectShape> directShapes)
        {
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
                            errorMessages[pipeOrDuct.Id].Add("Pipe/Duct không nằm hoàn toàn trong phạm vi được phép xuyên dầm.");
                            continue;
                        }

                        PlaceSleeveInstance(doc, sleeveSymbol, midpoint, direction, point1, point2, sleeveDiameter, sleevePlacements, entry.Key, pipeOrDuct, directShapes, errorMessages);
                    }
                }
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
            ScheduleFilter filter = new ScheduleFilter(schedule.Definition.GetField(0).FieldId, ScheduleFilterType.NotEqual, "");
            schedule.Definition.AddFilter(filter);
        }


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

        private DirectShape CreateDirectShapeFromFrameFace(Document doc, FrameObj frameObj, Face face)
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
            double widthMargin = frameHeight / 2; // Width margin
            double heightMargin = frameHeight / 4; // Height margin

            // Transform UV corners to XYZ using face.Evaluate method
            XYZ pt1 = planarFace.Evaluate(minUV);
            XYZ pt2 = planarFace.Evaluate(new UV(maxUV.U, minUV.V));
            XYZ pt3 = planarFace.Evaluate(maxUV);
            XYZ pt4 = planarFace.Evaluate(new UV(minUV.U, maxUV.V));

            //    // Xác định biên
            UV adjustedMin = new UV(minUV.U + widthMargin, minUV.V + heightMargin);
            UV adjustedMax = new UV(maxUV.U - widthMargin, maxUV.V - heightMargin);

            // Nếu điều chỉnh vượt ra ngoài biên, thì đặt lại giá trị hợp lý cho U, V
            if (adjustedMin.V >= adjustedMax.V)
            {
                adjustedMin = new UV(adjustedMin.U, minUV.V + (maxUV.V - minUV.V) / 4);
                adjustedMax = new UV(adjustedMax.U, maxUV.V - (maxUV.V - minUV.V) / 4);
            }
            if (adjustedMin.U >= adjustedMax.U)
            {
                adjustedMin = new UV(minUV.U + (maxUV.U - minUV.U) / 4, adjustedMin.V);
                adjustedMax = new UV(maxUV.U - (maxUV.U - minUV.U) / 4, adjustedMax.V);
            }

            // Tạo profile cho face bằng cách sử dụng các giá trị UV mới
            List<Curve> profile = new List<Curve> 
            {
                Line.CreateBound(face.Evaluate(adjustedMin), face.Evaluate(new UV(adjustedMin.U, adjustedMax.V))),
                Line.CreateBound(face.Evaluate(new UV(adjustedMin.U, adjustedMax.V)), face.Evaluate(adjustedMax)),
                Line.CreateBound(face.Evaluate(adjustedMax), face.Evaluate(new UV(adjustedMax.U, adjustedMin.V))),
                Line.CreateBound(face.Evaluate(new UV(adjustedMax.U, adjustedMin.V)), face.Evaluate(adjustedMin))
            };

           
            CurveLoop newCurveLoop = CurveLoop.Create(profile);

            // Extrusion direction typically normal to the face
            XYZ extrusionDirection = planarFace.FaceNormal;
            Solid directShapeSolid = GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { newCurveLoop }, extrusionDirection, 10.0 / 304.8);

            if (directShapeSolid != null && directShapeSolid.Volume > 0)
            {
                DirectShape directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                directShape.SetShape(new GeometryObject[] { directShapeSolid });
                directShape.SetName("Beam Intersection Zone");
                return directShape;
            }

            return null;
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
