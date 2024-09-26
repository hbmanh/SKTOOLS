//using System;
//using System.Collections.Generic;
//using System.Linq;
//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.DB.Structure;
//using Autodesk.Revit.UI;
//using SKRevitAddins.Utils;
//using UnitUtils = Autodesk.Revit.DB.UnitUtils;

//namespace SKRevitAddins.Commands.IntersectWithFrame
//{
//    [Transaction(TransactionMode.Manual)]
//    public class IntersectWithFrameCmd2 : IExternalCommand
//    {
//        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//        {
//            UIApplication uiapp = commandData.Application;
//            UIDocument uidoc = uiapp.ActiveUIDocument;
//            Document doc = uidoc.Document;
//            View activeView = uidoc.ActiveView;

//            var linkedDocs = GetLinkedDocuments(doc);
//            var pipesAndDucts = GetElementsOfType<MEPCurve>(doc).Where(e => e.Document.ActiveView.Id == activeView.Id).ToList();
//            var structuralFramings = new List<Element>();
//            foreach (var linkedDoc in linkedDocs)
//            {
//                structuralFramings.AddRange(GetElementsOfType<FamilyInstance>(linkedDoc, BuiltInCategory.OST_StructuralFraming));
//            }

//            var intersectionData = new Dictionary<ElementId, List<XYZ>>();
//            var errorMessages = new Dictionary<ElementId, HashSet<string>>();
//            var sleevePlacements = new Dictionary<ElementId, List<(XYZ, double)>>();
//            var directShapes = new List<DirectShape>();

//            ProcessIntersections(doc, structuralFramings, pipesAndDucts, intersectionData, directShapes);

//            var sleeveSymbol = GetSleeveSymbol(doc, ref message);
//            if (sleeveSymbol == null)
//                return Result.Failed;

//            PlaceSleeves(doc, sleeveSymbol, intersectionData, structuralFramings, sleevePlacements, errorMessages, directShapes);

//            if (errorMessages.Any())
//                CreateErrorSchedules(doc, errorMessages);

//            TaskDialog.Show("Intersections", $"Placed {intersectionData.Count} スリーブ_SK instances at intersections.");

//            ApplyFilterToDirectShapes(doc, activeView, directShapes);

//            return Result.Succeeded;
//        }

//        private List<Document> GetLinkedDocuments(Document doc)
//        {
//            return new FilteredElementCollector(doc)
//                .OfClass(typeof(RevitLinkInstance))
//                .Cast<RevitLinkInstance>()
//                .Select(link => link.GetLinkDocument())
//                .Where(linkedDoc => linkedDoc != null)
//                .ToList();
//        }

//        private List<Element> GetElementsOfType<T>(Document doc, BuiltInCategory? category = null) where T : Element
//        {
//            var collector = new FilteredElementCollector(doc)
//                .OfClass(typeof(T))
//                .WhereElementIsNotElementType();

//            if (category.HasValue)
//                collector.OfCategory(category.Value);

//            return collector.ToElements().ToList();
//        }

//        // Process intersections between structural framings and pipes/ducts
//        private void ProcessIntersections(Document doc, List<Element> structuralFramings, List<Element> pipesAndDucts, Dictionary<ElementId, List<XYZ>> intersectionData, List<DirectShape> directShapes)
//        {
//            using (Transaction trans = new Transaction(doc, "Place Sleeves and Create Direct Shapes"))
//            {
//                trans.Start();

//                var solidsIntersection = new List<Solid>();
//                foreach (var framing in structuralFramings)
//                {
//                    var framingGeometry = framing.get_Geometry(new Options());
//                    if (framingGeometry == null)
//                        continue;

//                    List<Solid> solids = GetSolidsFromGeometry(framingGeometry);
//                    Solid solid = TSolidUtils.UnionSolidList(solids);
//                    solidsIntersection.Add(solid);
//                    var surroundingFaces = GetSurroundingFaces(solid);
//                    foreach (Face face in surroundingFaces)
//                    {
//                        var directShape = CreateDirectShapeForBeamFace(doc, solid, face);
//                        if (directShape != null)
//                        {
//                            directShapes.Add(directShape);
//                        }

//                        foreach (var pipeOrDuct in pipesAndDucts)
//                        {
//                            var pipeOrDuctCurve = (pipeOrDuct.Location as LocationCurve)?.Curve;
//                            if (pipeOrDuctCurve == null)
//                                continue;

//                            if (face.Intersect(pipeOrDuctCurve, out IntersectionResultArray resultArray) == SetComparisonResult.Overlap)
//                            {
//                                if (!intersectionData.ContainsKey(pipeOrDuct.Id))
//                                {
//                                    intersectionData[pipeOrDuct.Id] = new List<XYZ>();
//                                }

//                                foreach (IntersectionResult intersectionResult in resultArray)
//                                {
//                                    intersectionData[pipeOrDuct.Id].Add(intersectionResult.XYZPoint);
//                                }
//                            }
//                        }
//                    }
//                }
//                //var unionSolid = solidsToUnion.UnionSolidList();
//                //unionSolid.BakeSolidToDirectShape(doc);
//                //var surroundingFaces = GetSurroundingFaces(unionSolid);
//                //foreach (Face face in surroundingFaces)
//                //{
//                //    var directShape = CreateDirectShapeForBeamFace(doc, unionSolid, face);
//                //    if (directShape != null)
//                //    {
//                //        directShapes.Add(directShape);
//                //    }

//                //    foreach (var pipeOrDuct in pipesAndDucts)
//                //    {
//                //        var pipeOrDuctCurve = (pipeOrDuct.Location as LocationCurve)?.Curve;
//                //        if (pipeOrDuctCurve == null)
//                //            continue;

//                //        if (face.Intersect(pipeOrDuctCurve, out IntersectionResultArray resultArray) == SetComparisonResult.Overlap)
//                //        {
//                //            if (!intersectionData.ContainsKey(pipeOrDuct.Id))
//                //            {
//                //                intersectionData[pipeOrDuct.Id] = new List<XYZ>();
//                //            }

//                //            foreach (IntersectionResult intersectionResult in resultArray)
//                //            {
//                //                intersectionData[pipeOrDuct.Id].Add(intersectionResult.XYZPoint);
//                //            }
//                //        }
//                //    }
//                //}
//                trans.Commit();
//            }
//        }

//        // Retrieve sleeve family symbol
//        private FamilySymbol GetSleeveSymbol(Document doc, ref string message)
//        {
//            var sleeveSymbol = new FilteredElementCollector(doc)
//                .OfCategory(BuiltInCategory.OST_PipeAccessory)
//                .OfClass(typeof(FamilySymbol))
//                .WhereElementIsElementType()
//                .Cast<FamilySymbol>()
//                .FirstOrDefault(symbol => symbol.FamilyName == "スリーブ_SK");

//            if (sleeveSymbol == null)
//            {
//                message = "The スリーブ_SK family could not be found.";
//                return null;
//            }

//            if (!sleeveSymbol.IsActive)
//            {
//                sleeveSymbol.Activate();
//                doc.Regenerate();
//            }

//            return sleeveSymbol;
//        }

//        // Place sleeves at intersection points
//        private void PlaceSleeves(Document doc, FamilySymbol sleeveSymbol, Dictionary<ElementId, List<XYZ>> intersectionData, List<Element> structuralFramings, Dictionary<ElementId, List<(XYZ, double)>> sleevePlacements, Dictionary<ElementId, HashSet<string>> errorMessages, List<DirectShape> directShapes)
//        {
//            using (Transaction trans = new Transaction(doc, "Place Sleeves"))
//            {
//                trans.Start();

//                foreach (var entry in intersectionData)
//                {
//                    var pipeOrDuct = doc.GetElement(entry.Key);
//                    var points = entry.Value;

//                    for (int i = 0; i < points.Count; i += 2)
//                    {
//                        if (i + 1 < points.Count)
//                        {
//                            XYZ point1 = points[i];
//                            XYZ point2 = points[i + 1];
//                            XYZ midpoint = (point1 + point2) / 2;
//                            XYZ direction = (point2 - point1).Normalize();

//                            double pipeDiameter = pipeOrDuct.LookupParameter("Diameter")?.AsDouble() ?? 0;
//                            double sleeveDiameter = pipeDiameter + UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Millimeters);
//                            double beamHeight = GetBeamHeight(point1, structuralFramings);

//                            HashSet<string> errors = ValidateSleevePlacement(sleeveDiameter, beamHeight, midpoint, sleevePlacements, entry.Key);

//                            if (errors.Count > 0)
//                            {
//                                if (!errorMessages.ContainsKey(pipeOrDuct.Id))
//                                {
//                                    errorMessages[pipeOrDuct.Id] = new HashSet<string>();
//                                }
//                                errorMessages[pipeOrDuct.Id].UnionWith(errors);
//                                continue;
//                            }

//                            // Kiểm tra nếu pipe/duct không nằm hoàn toàn trong phạm vi direct shape sẽ không được tạo ra
//                            if (!IsPointWithinAnyDirectShape(point1, directShapes) || !IsPointWithinAnyDirectShape(point2, directShapes))
//                            {
//                                if (!errorMessages.ContainsKey(pipeOrDuct.Id))
//                                {
//                                    errorMessages[pipeOrDuct.Id] = new HashSet<string>();
//                                }
//                                errorMessages[pipeOrDuct.Id].Add("Pipe/Duct does not fully lie within the permissible direct shape boundaries.");
//                                continue;
//                            }

//                            PlaceSleeveInstance(doc, sleeveSymbol, midpoint, direction, point1, point2, sleeveDiameter, sleevePlacements, entry.Key, pipeOrDuct, directShapes, errorMessages);
//                        }
//                    }
//                }

//                trans.Commit();
//            }
//        }

//        // Validate sleeve placement conditions
//        private HashSet<string> ValidateSleevePlacement(double sleeveDiameter, double beamHeight, XYZ midpoint, Dictionary<ElementId, List<(XYZ, double)>> sleevePlacements, ElementId pipeOrDuctId)
//        {
//            HashSet<string> errors = new HashSet<string>();

//            if (sleeveDiameter > UnitUtils.ConvertToInternalUnits(750, UnitTypeId.Millimeters))
//            {
//                errors.Add("OD > 750mm");
//            }

//            if (sleeveDiameter > beamHeight / 3)
//            {
//                errors.Add("OD > H/3");
//            }

//            if (!sleevePlacements.ContainsKey(pipeOrDuctId))
//            {
//                sleevePlacements[pipeOrDuctId] = new List<(XYZ, double)>();
//            }

//            foreach (var (otherMidpoint, otherDiameter) in sleevePlacements[pipeOrDuctId])
//            {
//                double minDistance = (sleeveDiameter + otherDiameter) * 1.5;
//                if (Math.Abs(midpoint.X - otherMidpoint.X) < minDistance || Math.Abs(midpoint.Y - otherMidpoint.Y) < minDistance)
//                {
//                    errors.Add($"Distance between sleeves < (OD1 + OD2)*3/2");
//                    break;
//                }
//            }

//            return errors;
//        }

//        // Place sleeve instance at midpoint
//        private void PlaceSleeveInstance(Document doc, FamilySymbol sleeveSymbol, XYZ midpoint, XYZ direction, XYZ point1, XYZ point2, double sleeveDiameter, Dictionary<ElementId, List<(XYZ, double)>> sleevePlacements, ElementId entryKey, Element pipeOrDuct, List<DirectShape> directShapes, Dictionary<ElementId, HashSet<string>> errorMessages)
//        {
//            FamilyInstance sleeveInstance = doc.Create.NewFamilyInstance(midpoint, sleeveSymbol, StructuralType.NonStructural);

//            Line axis = Line.CreateBound(midpoint, midpoint + XYZ.BasisZ);
//            double angle = XYZ.BasisX.AngleTo(direction);
//            double additionalRotation = Math.PI / 2;
//            ElementTransformUtils.RotateElement(doc, sleeveInstance.Id, axis, angle + additionalRotation);

//            Parameter lengthParam = sleeveInstance.LookupParameter("L");
//            if (lengthParam != null)
//            {
//                lengthParam.Set(point1.DistanceTo(point2));
//            }

//            Parameter odParam = sleeveInstance.LookupParameter("OD");
//            if (odParam != null)
//            {
//                odParam.Set(sleeveDiameter);
//            }

//            sleevePlacements[entryKey].Add((midpoint, sleeveDiameter));

//            bool isWithinDirectShape = directShapes.Any(ds => IsPointWithinDirectShape(midpoint, ds));
//            if (!isWithinDirectShape)
//            {
//                if (!errorMessages.ContainsKey(pipeOrDuct.Id))
//                {
//                    errorMessages[pipeOrDuct.Id] = new HashSet<string>();
//                }
//                errorMessages[pipeOrDuct.Id].Add("Sleeve Placed Outside Permissible Beam Penetration Range.");
//            }
//        }

//        // Handle errors and prompt to save error messages
//        private void CreateErrorSchedules(Document doc, Dictionary<ElementId, HashSet<string>> errorMessages)
//        {
//            // Create a schedule for pipes
//            CreateErrorSchedule(doc, errorMessages, "PipeErrorSchedule", BuiltInCategory.OST_PipeCurves);

//            // Create a schedule for ducts
//            CreateErrorSchedule(doc, errorMessages, "DuctErrorSchedule", BuiltInCategory.OST_DuctCurves);
//        }

//        private void CreateErrorSchedule(Document doc, Dictionary<ElementId, HashSet<string>> errorMessages, string scheduleName, BuiltInCategory category)
//        {
//            using (Transaction trans = new Transaction(doc, "Create Error Schedule"))
//            {
//                trans.Start();

//                ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, new ElementId(category));

//                schedule.Name = scheduleName;

//                SchedulableField markField = schedule.Definition.GetSchedulableFields().First(sf => sf.GetName(doc) == "Mark");
//                SchedulableField commentField = schedule.Definition.GetSchedulableFields().First(sf => sf.GetName(doc) == "Comments");

//                schedule.Definition.AddField(ScheduleFieldType.Instance, markField.ParameterId);
//                schedule.Definition.AddField(ScheduleFieldType.Instance, commentField.ParameterId);

//                int markIndex = 1;
//                foreach (var kvp in errorMessages)
//                {
//                    Element element = doc.GetElement(kvp.Key);
//                    if (element.Category.Id.IntegerValue != (int)category)
//                        continue;

//                    Parameter markParam = element.LookupParameter("Mark");
//                    if (markParam != null && kvp.Value.Any())
//                    {
//                        markParam.Set(markIndex.ToString());
//                        markIndex++;
//                    }

//                    Parameter commentParam = element.LookupParameter("Comments");
//                    if (commentParam != null)
//                    {
//                        commentParam.Set(string.Join(", ", kvp.Value));
//                    }
//                }

//                // Apply filter to remove elements without errors
//                ScheduleFilter filter = new ScheduleFilter(schedule.Definition.GetField(0).FieldId, ScheduleFilterType.Equal, "");
//                schedule.Definition.AddFilter(filter);

//                trans.Commit();
//            }
//        }

//        // Get beam height from structural framings
//        private double GetBeamHeight(XYZ point, List<Element> structuralFramings)
//        {
//            foreach (var framing in structuralFramings)
//            {
//                if (IsPointOnElement(framing, point))
//                {
//                    BoundingBoxXYZ boundingBox = framing.get_BoundingBox(null);
//                    if (boundingBox != null)
//                    {
//                        return boundingBox.Max.Z - boundingBox.Min.Z;
//                    }
//                }
//            }
//            return 0;
//        }

//        // Check if a point is on a given element
//        private bool IsPointOnElement(Element element, XYZ point)
//        {
//            var boundingBox = element.get_BoundingBox(null);
//            return boundingBox != null &&
//                   point.X >= boundingBox.Min.X && point.X <= boundingBox.Max.X &&
//                   point.Y >= boundingBox.Min.Y && point.Y <= boundingBox.Max.Y &&
//                   point.Z >= boundingBox.Min.Z && point.Z <= boundingBox.Max.Z;
//        }

//        // Get solids from geometry element
//        private List<Solid> GetSolidsFromGeometry(GeometryElement geometryElement)
//        {
//            List<Solid> solids = new List<Solid>();

//            foreach (GeometryObject geomObj in geometryElement)
//            {
//                if (geomObj is Solid solid && solid.Volume > 0)
//                {
//                    solids.Add(solid);
//                }
//                else if (geomObj is GeometryInstance geomInstance)
//                {
//                    GeometryElement instanceGeometry = geomInstance.GetInstanceGeometry();
//                    solids.AddRange(GetSolidsFromGeometry(instanceGeometry));
//                }
//            }

//            return solids;
//        }

//        private List<Face> GetSurroundingFaces(Solid solid)
//        {
//            var faces = solid.GetSolidVerticalFaces();

//            var faceAreas = faces.Select(face => new { Face = face, Area = GetFaceArea(face) }).ToList();

//            var sortedFaceAreas = faceAreas.OrderBy(f => f.Area).ToList();

//            var remainingFaces = sortedFaceAreas.Skip(2).Select(f => f.Face).ToList();

//            return remainingFaces;
//        }

//        // Get area of a face
//        private double GetFaceArea(Face face)
//        {
//            Mesh mesh = face.Triangulate();
//            double area = 0;
//            int numTriangles = mesh.NumTriangles;

//            for (int i = 0; i < numTriangles; i++)
//            {
//                MeshTriangle triangle = mesh.get_Triangle(i);
//                XYZ p0 = triangle.get_Vertex(0);
//                XYZ p1 = triangle.get_Vertex(1);
//                XYZ p2 = triangle.get_Vertex(2);
//                area += 0.5 * ((p1 - p0).CrossProduct(p2 - p0)).GetLength();
//            }

//            return area;
//        }

//        // Create direct shape for a beam face
//        private List<DirectShape> CreateDirectShapeForBeamFace(Document doc, Solid solid, Face face, double beamHeight)
//        {
//            List<DirectShape> createdDirectShapes = new List<DirectShape>();

//            BoundingBoxUV boundingBox = face.GetBoundingBox();
//            double margin = beamHeight / 2;
//            UV min = boundingBox.Min;
//            UV max = boundingBox.Max;

//            UV adjustedMin = new UV(min.U + margin, min.V + margin);
//            UV adjustedMax = new UV(max.U - margin, max.V - margin);

//            double uLength = adjustedMax.U - adjustedMin.U;
//            double vLength = adjustedMax.V - adjustedMin.V;

//            if (uLength <= 0 || vLength <= 0) // Check if the remaining area is valid
//                return createdDirectShapes;

//            // Calculate number of segments based on the margin
//            int numSegments = (int)Math.Floor(uLength / margin);

//            for (int i = 0; i < numSegments; i++)
//            {
//                UV segmentMin = new UV(adjustedMin.U + i * margin, adjustedMin.V);
//                UV segmentMax = new UV(segmentMin.U + margin, adjustedMax.V);

//                List<Curve> profile = new List<Curve>();
//                XYZ p1 = face.Evaluate(segmentMin);
//                XYZ p2 = face.Evaluate(new UV(segmentMin.U, segmentMax.V));
//                XYZ p3 = face.Evaluate(segmentMax);
//                XYZ p4 = face.Evaluate(new UV(segmentMax.U, segmentMin.V));

//                profile.Add(Line.CreateBound(p1, p2));
//                profile.Add(Line.CreateBound(p2, p3));
//                profile.Add(Line.CreateBound(p3, p4));
//                profile.Add(Line.CreateBound(p4, p1));

//                CurveLoop curveLoop = CurveLoop.Create(profile);
//                List<CurveLoop> curveLoops = new List<CurveLoop> { curveLoop };

//                XYZ extrusionDirection = face.ComputeNormal(UV.Zero);
//                double extrusionDepth = 10.0 / 304.8; // Convert mm to internal units (feet)

//                Solid directShapeSolid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoops, extrusionDirection, extrusionDepth);
//                DirectShape directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
//                directShape.SetShape(new GeometryObject[] { directShapeSolid });
//                directShape.SetName("Beam Intersection Zone Segment " + (i + 1));

//                createdDirectShapes.Add(directShape);
//            }

//            return createdDirectShapes;
//        }


//        // Check if a point is within a direct shape
//        private bool IsPointWithinDirectShape(XYZ point, DirectShape directShape)
//        {
//            var boundingBox = directShape.get_BoundingBox(null);
//            if (boundingBox == null)
//                return false;

//            return (point.X >= boundingBox.Min.X && point.X <= boundingBox.Max.X) &&
//                   (point.Y >= boundingBox.Min.Y && point.Y <= boundingBox.Max.Y) &&
//                   (point.Z >= boundingBox.Min.Z && point.Z <= boundingBox.Max.Z);
//        }

//        // Check if a point is within any of the direct shapes
//        private bool IsPointWithinAnyDirectShape(XYZ point, List<DirectShape> directShapes)
//        {
//            return directShapes.Any(ds => IsPointWithinDirectShape(point, ds));
//        }

//        // Apply filter to direct shapes and set color
//        private void ApplyFilterToDirectShapes(Document doc, View view, List<DirectShape> directShapes)
//        {
//            using (Transaction trans = new Transaction(doc, "Apply Filter to Direct Shapes"))
//            {
//                trans.Start();

//                ParameterFilterElement parameterFilter = new FilteredElementCollector(doc)
//                    .OfClass(typeof(ParameterFilterElement))
//                    .Cast<ParameterFilterElement>()
//                    .FirstOrDefault(f => f.Name == "DirectShape Filter");

//                if (parameterFilter == null)
//                {
//                    ElementCategoryFilter directShapeFilter = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel);
//                    parameterFilter = ParameterFilterElement.Create(doc, "DirectShape Filter", new List<ElementId> { new ElementId(BuiltInCategory.OST_GenericModel) });
//                    view.AddFilter(parameterFilter.Id); 
//                }
//                else if (!view.IsFilterApplied(parameterFilter.Id)) 
//                {
//                    view.AddFilter(parameterFilter.Id); 
//                }

//                ElementId solidFillPatternId = null;
//                List<FillPatternElement> fillPatternList = new FilteredElementCollector(doc)
//                    .WherePasses(new ElementClassFilter(typeof(FillPatternElement))).
//                    ToElements().Cast<FillPatternElement>().ToList();

//                foreach (FillPatternElement fp in fillPatternList)
//                {
//                    if (fp.GetFillPattern().IsSolidFill)
//                    {
//                        solidFillPatternId = fp.Id;
//                        break;
//                    }
//                }

//                OverrideGraphicSettings overrideSettings = new OverrideGraphicSettings();
//                overrideSettings.SetSurfaceForegroundPatternColor(new Color(0, 255, 0));
//                overrideSettings.SetSurfaceForegroundPatternId(solidFillPatternId);

//                foreach (var directShape in directShapes)
//                {
//                    view.SetElementOverrides(directShape.Id, overrideSettings);
//                }

//                trans.Commit();
//            }
//        }


//    }
//}
