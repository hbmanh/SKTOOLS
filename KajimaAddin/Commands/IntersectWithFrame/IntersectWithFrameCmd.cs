using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SKToolsAddins.Commands.IntersectWithFrame
{
    [Transaction(TransactionMode.Manual)]
    public class IntersectWithFrameCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var linkedDocs = GetLinkedDocuments(doc);
            var pipesAndDucts = GetElementsOfType<MEPCurve>(doc);
            var structuralFramings = new List<Element>();
            foreach (var linkedDoc in linkedDocs)
            {
                structuralFramings.AddRange(GetElementsOfType<FamilyInstance>(linkedDoc, BuiltInCategory.OST_StructuralFraming));
            }

            var intersectionData = new Dictionary<ElementId, List<XYZ>>();
            var errorMessages = new Dictionary<ElementId, HashSet<string>>();
            var sleevePlacements = new Dictionary<ElementId, List<(XYZ, double)>>();
            var directShapes = new List<DirectShape>();

            ProcessIntersections(doc, structuralFramings, pipesAndDucts, intersectionData, directShapes);

            var sleeveSymbol = GetSleeveSymbol(doc, ref message);
            if (sleeveSymbol == null)
                return Result.Failed;

            PlaceSleeves(doc, sleeveSymbol, intersectionData, structuralFramings, sleevePlacements, errorMessages, directShapes);

            if (errorMessages.Any())
                HandleErrors(errorMessages);
            else
                TaskDialog.Show("Intersections", $"Placed {intersectionData.Count} スリーブ_SK instances at intersections.");

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

        private void ProcessIntersections(Document doc, List<Element> structuralFramings, List<Element> pipesAndDucts, Dictionary<ElementId, List<XYZ>> intersectionData, List<DirectShape> directShapes)
        {
            using (Transaction trans = new Transaction(doc, "Place Sleeves and Create Direct Shapes"))
            {
                trans.Start();
                foreach (var framing in structuralFramings)
                {
                    var framingGeometry = framing.get_Geometry(new Options());
                    if (framingGeometry == null)
                        continue;

                    List<Solid> solids = GetSolidsFromGeometry(framingGeometry);

                    foreach (Solid solid in solids)
                    {
                        var surroundingFaces = GetSurroundingFaces(solid);
                        foreach (Face face in surroundingFaces)
                        {
                            var directShape = CreateDirectShapeForBeamFace(doc, solid, face);
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

                            // Logic to check if the beam intersects and is joined with another beam
                            foreach (var otherFraming in structuralFramings)
                            {
                                if (otherFraming.Id != framing.Id && AreBeamsBoundingBoxesIntersect(framing, otherFraming))
                                {
                                    // Your logic to create additional direct shapes when beams are joined
                                    var otherFramingGeometry = otherFraming.get_Geometry(new Options());
                                    if (otherFramingGeometry == null)
                                        continue;

                                    List<Solid> otherSolids = GetSolidsFromGeometry(otherFramingGeometry);
                                    foreach (Solid otherSolid in otherSolids)
                                    {
                                        var otherSurroundingFaces = GetSurroundingFaces(otherSolid);
                                        foreach (Face otherFace in otherSurroundingFaces)
                                        {
                                            var additionalDirectShape = CreateDirectShapeForBeamFace(doc, otherSolid, otherFace);
                                            if (additionalDirectShape != null)
                                            {
                                                directShapes.Add(additionalDirectShape);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                trans.Commit();
            }
        }

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

                            PlaceSleeveInstance(doc, sleeveSymbol, midpoint, direction, point1, point2, sleeveDiameter, sleevePlacements, entry.Key, pipeOrDuct, directShapes, errorMessages);
                        }
                    }
                }

                trans.Commit();
            }
        }

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

        private void HandleErrors(Dictionary<ElementId, HashSet<string>> errorMessages)
        {
            string errorMsg = string.Join("\n", errorMessages.Select(kv => $"ID: {kv.Key} - Errors: {string.Join(", ", kv.Value.Distinct())}"));

            string previewErrorMsg = errorMsg.Length > 500 ? errorMsg.Substring(0, 500) + "..." : errorMsg;

            TaskDialog taskDialog = new TaskDialog("Invalid Pipes/Ducts")
            {
                MainContent = "There are errors in placing some sleeves. Do you want to save these errors to a text file?",
                ExpandedContent = previewErrorMsg,
                CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
            };

            if (taskDialog.Show() == TaskDialogResult.Yes)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                    Title = "Save Error Messages"
                };
                if (saveFileDialog.ShowDialog() == true)
                {
                    File.WriteAllText(saveFileDialog.FileName, errorMsg);
                    TaskDialog.Show("Invalid Pipes/Ducts", $"Errors have been written to {saveFileDialog.FileName}");
                }
            }
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

        private bool IsPointOnElement(Element element, XYZ point)
        {
            var boundingBox = element.get_BoundingBox(null);
            return boundingBox != null &&
                   point.X >= boundingBox.Min.X && point.X <= boundingBox.Max.X &&
                   point.Y >= boundingBox.Min.Y && point.Y <= boundingBox.Max.Y &&
                   point.Z >= boundingBox.Min.Z && point.Z <= boundingBox.Max.Z;
        }

        private List<Solid> GetSolidsFromGeometry(GeometryElement geometryElement)
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

        private List<Face> GetSurroundingFaces(Solid solid)
        {
            var faces = solid.Faces.Cast<Face>().Where(face => !IsTopOrBottomFace(face)).ToList();

            var faceAreas = faces.Select(face => new { Face = face, Area = GetFaceArea(face) }).ToList();

            var sortedFaceAreas = faceAreas.OrderBy(f => f.Area).ToList();

            sortedFaceAreas.RemoveAt(0);
            sortedFaceAreas.RemoveAt(0);

            return sortedFaceAreas.Select(f => f.Face).ToList();
        }

        private bool IsTopOrBottomFace(Face face)
        {
            XYZ normal = face.ComputeNormal(new UV(0.5, 0.5));
            return Math.Abs(normal.Z) > 0.9;
        }

        private double GetFaceArea(Face face)
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

        private DirectShape CreateDirectShapeForBeamFace(Document doc, Solid solid, Face face)
        {
            BoundingBoxUV boundingBox = face.GetBoundingBox();
            UV min = boundingBox.Min;
            UV max = boundingBox.Max;

            BoundingBoxXYZ solidBoundingBox = solid.GetBoundingBox();
            double beamHeight = solidBoundingBox.Max.Z - solidBoundingBox.Min.Z;

            double heightMargin = beamHeight / 4;
            double widthMargin = beamHeight;

            UV adjustedMin = new UV(min.U + widthMargin, min.V + widthMargin);
            UV adjustedMax = new UV(max.U - widthMargin, max.V - widthMargin);

            if (adjustedMin.V >= adjustedMax.V)
            {
                adjustedMin = new UV(adjustedMin.U, min.V + (max.V - min.V) / 4);
                adjustedMax = new UV(adjustedMax.U, max.V - (max.V - min.V) / 4);
            }

            if (adjustedMin.U >= adjustedMax.U)
            {
                adjustedMin = new UV(min.U + (max.U - min.U) / 4, adjustedMin.V);
                adjustedMax = new UV(max.U - (max.U - min.U) / 4, adjustedMax.V);
            }

            List<Curve> profile = new List<Curve>
            {
                Line.CreateBound(face.Evaluate(adjustedMin), face.Evaluate(new UV(adjustedMin.U, adjustedMax.V))),
                Line.CreateBound(face.Evaluate(new UV(adjustedMin.U, adjustedMax.V)), face.Evaluate(adjustedMax)),
                Line.CreateBound(face.Evaluate(adjustedMax), face.Evaluate(new UV(adjustedMax.U, adjustedMin.V))),
                Line.CreateBound(face.Evaluate(new UV(adjustedMax.U, adjustedMin.V)), face.Evaluate(adjustedMin))
            };

            CurveLoop curveLoop = CurveLoop.Create(profile);
            List<CurveLoop> curveLoops = new List<CurveLoop> { curveLoop };
            XYZ extrusionDirection = face.ComputeNormal(UV.Zero);
            Solid directShapeSolid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoops, extrusionDirection, 10.0 / 304.8);

            DirectShape directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
            directShape.SetShape(new GeometryObject[] { directShapeSolid });
            directShape.SetName("Beam Intersection Zone");

            return directShape;
        }

        private bool IsPointWithinDirectShape(XYZ point, DirectShape directShape)
        {
            var boundingBox = directShape.get_BoundingBox(null);
            if (boundingBox == null)
                return false;

            return (point.X >= boundingBox.Min.X && point.X <= boundingBox.Max.X) &&
                   (point.Y >= boundingBox.Min.Y && point.Y <= boundingBox.Max.Y) &&
                   (point.Z >= boundingBox.Min.Z && point.Z <= boundingBox.Max.Z);
        }

        private bool AreBeamsBoundingBoxesIntersect(Element beamA, Element beamB)
        {
            var bboxA = beamA.get_BoundingBox(null);
            var bboxB = beamB.get_BoundingBox(null);

            if (bboxA == null || bboxB == null)
                return false;

            return !(bboxA.Max.X < bboxB.Min.X || bboxA.Min.X > bboxB.Max.X ||
                     bboxA.Max.Y < bboxB.Min.Y || bboxA.Min.Y > bboxB.Max.Y ||
                     bboxA.Max.Z < bboxB.Min.Z || bboxA.Min.Z > bboxB.Max.Z);
        }
    }
}
