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
            View activeView = uidoc.ActiveView;

            var linkedDocs = GetLinkedDocuments(doc);
            var pipesAndDucts = GetElementsOfType<MEPCurve>(doc).Where(e => e.Document.ActiveView.Id == activeView.Id).ToList();
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

            ApplyFilterToDirectShapes(doc, activeView, directShapes);

            return Result.Succeeded;
        }

        // Retrieve linked documents
        private List<Document> GetLinkedDocuments(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Select(link => link.GetLinkDocument())
                .Where(linkedDoc => linkedDoc != null)
                .ToList();
        }

        // Retrieve elements of specified type and category
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
                        }
                    }
                }
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

                            // Kiểm tra xem vị trí midpoint đã có sleeve nào chưa
                            if (!sleevePlacements.ContainsKey(entry.Key))
                            {
                                sleevePlacements[entry.Key] = new List<(XYZ, double)>();
                            }

                            bool placementValid = true;
                            foreach (var (otherMidpoint, otherDiameter) in sleevePlacements[entry.Key])
                            {
                                double minDistance = (sleeveDiameter + otherDiameter) * 1.5;
                                if (midpoint.DistanceTo(otherMidpoint) < minDistance)
                                {
                                    placementValid = false;
                                    break;
                                }
                            }

                            if (placementValid)
                            {
                                PlaceSleeveInstance(doc, sleeveSymbol, midpoint, direction, point1, point2, sleeveDiameter, sleevePlacements, entry.Key, pipeOrDuct, directShapes, errorMessages);
                            }
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
                if (midpoint.DistanceTo(otherMidpoint) < minDistance)
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

        // Get surrounding faces excluding top and bottom faces
        private List<Face> GetSurroundingFaces(Solid solid)
        {
            var faces = solid.Faces.Cast<Face>().Where(face => !IsTopOrBottomFace(face)).ToList();

            var faceAreas = faces.Select(face => new { Face = face, Area = GetFaceArea(face) }).ToList();

            var sortedFaceAreas = faceAreas.OrderBy(f => f.Area).ToList();

            sortedFaceAreas.RemoveAt(0);
            sortedFaceAreas.RemoveAt(0);

            return sortedFaceAreas.Select(f => f.Face).ToList();
        }

        // Check if a face is the top or bottom face
        private bool IsTopOrBottomFace(Face face)
        {
            XYZ normal = face.ComputeNormal(new UV(0.5, 0.5));
            return Math.Abs(normal.Z) > 0.9; // Assuming Z direction is vertical
        }

        // Get area of a face
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

        // Create direct shape for a beam face
        private DirectShape CreateDirectShapeForBeamFace(Document doc, Solid solid, Face face)
        {
            // Lấy BoundingBoxUV của mặt
            BoundingBoxUV boundingBox = face.GetBoundingBox();
            UV min = boundingBox.Min;
            UV max = boundingBox.Max;

            // Lấy BoundingBoxXYZ của khối solid để tính chiều cao của dầm
            BoundingBoxXYZ solidBoundingBox = solid.GetBoundingBox();
            double beamHeight = solidBoundingBox.Max.Z - solidBoundingBox.Min.Z;

            // Xác định khoảng cách biên cần thiết từ top và bottom của dầm
            double heightMargin = beamHeight / 4;
            double widthMargin = beamHeight;

            // Điều chỉnh các giá trị min và max của BoundingBoxUV để tạo khoảng cách biên cần thiết
            UV adjustedMin = new UV(min.U + widthMargin, min.V + widthMargin);
            UV adjustedMax = new UV(max.U - widthMargin, max.V - widthMargin);

            // Nếu điều chỉnh vượt ra ngoài biên, thì đặt lại giá trị hợp lý cho V
            if (adjustedMin.V >= adjustedMax.V)
            {
                adjustedMin = new UV(adjustedMin.U, min.V + (max.V - min.V) / 4);
                adjustedMax = new UV(adjustedMax.U, max.V - (max.V - min.V) / 4);
            }

            // Nếu điều chỉnh vượt ra ngoài biên, thì đặt lại giá trị hợp lý cho U
            if (adjustedMin.U >= adjustedMax.U)
            {
                adjustedMin = new UV(min.U + (max.U - min.U) / 4, adjustedMin.V);
                adjustedMax = new UV(max.U - (max.U - min.U) / 4, adjustedMax.V);
            }

            // Tạo các đường bao quanh (profile) cho mặt bằng cách sử dụng các giá trị UV đã điều chỉnh
            List<Curve> profile = new List<Curve>
            {
                Line.CreateBound(face.Evaluate(adjustedMin), face.Evaluate(new UV(adjustedMin.U, adjustedMax.V))),
                Line.CreateBound(face.Evaluate(new UV(adjustedMin.U, adjustedMax.V)), face.Evaluate(adjustedMax)),
                Line.CreateBound(face.Evaluate(adjustedMax), face.Evaluate(new UV(adjustedMax.U, adjustedMin.V))),
                Line.CreateBound(face.Evaluate(new UV(adjustedMax.U, adjustedMin.V)), face.Evaluate(adjustedMin))
            };

            // Tạo CurveLoop từ profile
            CurveLoop curveLoop = CurveLoop.Create(profile);
            List<CurveLoop> curveLoops = new List<CurveLoop> { curveLoop };

            // Tính toán hướng đùn (extrusion direction) dựa trên pháp tuyến của mặt
            XYZ extrusionDirection = face.ComputeNormal(UV.Zero);

            // Tạo khối solid cho DirectShape bằng cách đùn CurveLoop
            Solid directShapeSolid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoops, extrusionDirection, 10.0 / 304.8); // 10mm chuyển đổi sang đơn vị nội bộ

            // Tạo DirectShape trong tài liệu Revit
            DirectShape directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
            directShape.SetShape(new GeometryObject[] { directShapeSolid });
            directShape.SetName("Beam Intersection Zone");

            return directShape;
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

        // Apply filter to direct shapes and set color to green
        private void ApplyFilterToDirectShapes(Document doc, View view, List<DirectShape> directShapes)
        {
            using (Transaction trans = new Transaction(doc, "Apply Filter to Direct Shapes"))
            {
                trans.Start();

                // Create a filter for direct shapes
                ElementCategoryFilter directShapeFilter = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel);
                ParameterFilterElement parameterFilter = ParameterFilterElement.Create(doc, "DirectShape Filter", new List<ElementId> { new ElementId(BuiltInCategory.OST_GenericModel) });

                // Apply the filter to the active view
                view.AddFilter(parameterFilter.Id);
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

                // Set the color to green for the direct shapes
                OverrideGraphicSettings overrideSettings = new OverrideGraphicSettings();
                view.SetFilterOverrides(parameterFilter.Id, overrideSettings.SetSurfaceForegroundPatternColor(new Color(0, 255, 0)));
                view.SetFilterOverrides(parameterFilter.Id, overrideSettings.SetSurfaceForegroundPatternId(solidFillPatternId));


                foreach (var directShape in directShapes)
                {
                    view.SetElementOverrides(directShape.Id, overrideSettings);
                }

                trans.Commit();
            }
        }
    }
}
