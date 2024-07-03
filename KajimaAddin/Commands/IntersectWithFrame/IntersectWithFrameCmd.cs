using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Microsoft.Win32; // Namespace for SaveFileDialog
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

            // Get all linked documents
            var linkedDocs = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Select(link => link.GetLinkDocument())
                .Where(linkedDoc => linkedDoc != null)
                .ToList();

            // Get all pipes and ducts in the current document
            var pipesAndDucts = new FilteredElementCollector(doc)
                .OfClass(typeof(MEPCurve))
                .WhereElementIsNotElementType()
                .ToElements();

            // Get all structural framings in the linked documents
            var structuralFramings = linkedDocs
                .SelectMany(linkedDoc => new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .ToElements())
                .ToList();

            // Dictionary to store intersection results with midpoint and direction
            var intersectionData = new Dictionary<ElementId, List<XYZ>>();
            var errorMessages = new Dictionary<ElementId, HashSet<string>>();
            var sleevePlacements = new Dictionary<ElementId, List<(XYZ, double)>>(); // Track placed sleeves with their diameters
            var directShapes = new List<DirectShape>();

            using (var trans = new Transaction(doc, "Place Sleeves and Create Direct Shapes"))
            {
                trans.Start();

                foreach (var pipeOrDuct in pipesAndDucts)
                {
                    var pipeOrDuctCurve = (pipeOrDuct.Location as LocationCurve)?.Curve;
                    if (pipeOrDuctCurve == null) continue;

                    foreach (var framing in structuralFramings)
                    {
                        var framingGeometry = framing.get_Geometry(new Options());
                        if (framingGeometry == null) continue;

                        var solids = GetSolidsFromGeometry(framingGeometry);

                        foreach (var solid in solids)
                        {
                            foreach (Face face in solid.Faces)
                            {
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

                                    var directShape = CreateDirectShapeForBeamFace(doc, solid, face);
                                    if (directShape != null)
                                    {
                                        directShapes.Add(directShape);
                                    }
                                }
                            }
                        }
                    }
                }

                // Load the スリーブ_SK family symbol
                var sleeveSymbol = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeAccessory)
                    .OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(symbol => symbol.FamilyName == "スリーブ_SK");

                if (sleeveSymbol == null)
                {
                    message = "The スリーブ_SK family could not be found.";
                    return Result.Failed;
                }

                if (!sleeveSymbol.IsActive)
                {
                    sleeveSymbol.Activate();
                    doc.Regenerate();
                }

                foreach (var entry in intersectionData)
                {
                    var pipeOrDuct = doc.GetElement(entry.Key);
                    var points = entry.Value;

                    for (int i = 0; i < points.Count - 1; i += 2)
                    {
                        var point1 = points[i];
                        var point2 = points[i + 1];
                        var midpoint = (point1 + point2) / 2;
                        var direction = (point2 - point1).Normalize();

                        var pipeDiameter = pipeOrDuct.LookupParameter("Diameter")?.AsDouble() ?? 0;
                        var sleeveDiameter = pipeDiameter + UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Millimeters); // Adding 50mm and converting to feet
                        var beamHeight = GetBeamHeight(point1, point2, structuralFramings);

                        var errors = new HashSet<string>();

                        if (sleeveDiameter > UnitUtils.ConvertToInternalUnits(750, UnitTypeId.Millimeters))
                        {
                            errors.Add("OD > 750mm");
                        }

                        if (sleeveDiameter > beamHeight / 3)
                        {
                            errors.Add("OD > H/3");
                        }

                        if (!sleevePlacements.ContainsKey(entry.Key))
                        {
                            sleevePlacements[entry.Key] = new List<(XYZ, double)>();
                        }

                        // Check distance to other sleeves in the same beam
                        foreach (var (otherMidpoint, otherDiameter) in sleevePlacements[entry.Key])
                        {
                            var minDistance = (sleeveDiameter + otherDiameter) * 1.5;
                            if (Math.Abs(midpoint.X - otherMidpoint.X) < minDistance || Math.Abs(midpoint.Y - otherMidpoint.Y) < minDistance)
                            {
                                errors.Add("Distance between sleeves < (OD1 + OD2)*3/2");
                                break;
                            }
                        }

                        if (errors.Any())
                        {
                            if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                            {
                                errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                            }
                            errorMessages[pipeOrDuct.Id].UnionWith(errors);
                            continue;
                        }

                        // Place the sleeve instance
                        var sleeveInstance = doc.Create.NewFamilyInstance(midpoint, sleeveSymbol, StructuralType.NonStructural);

                        // Rotate the sleeve to be parallel with the direction vector plus an additional 90 degrees
                        var axis = Line.CreateBound(midpoint, midpoint + XYZ.BasisZ);
                        var angle = XYZ.BasisX.AngleTo(direction);
                        var additionalRotation = Math.PI / 2; // 90 degrees in radians
                        ElementTransformUtils.RotateElement(doc, sleeveInstance.Id, axis, angle + additionalRotation);

                        // Set the parameter L to the distance between the intersection points
                        var lengthParam = sleeveInstance.LookupParameter("L");
                        lengthParam?.Set(point1.DistanceTo(point2));

                        // Set the parameter OD to the calculated sleeve diameter
                        var odParam = sleeveInstance.LookupParameter("OD");
                        odParam?.Set(sleeveDiameter);

                        // Track the sleeve placement
                        sleevePlacements[entry.Key].Add((midpoint, sleeveDiameter));

                        // Check if the midpoint of the sleeve is within any direct shape
                        if (!directShapes.Any(ds => IsPointWithinDirectShape(midpoint, ds)))
                        {
                            if (!errorMessages.ContainsKey(pipeOrDuct.Id))
                            {
                                errorMessages[pipeOrDuct.Id] = new HashSet<string>();
                            }
                            errorMessages[pipeOrDuct.Id].Add("Sleeve Placed Outside Permissible Beam Penetration Range.");
                        }
                    }
                }

                trans.Commit();
            }

            if (errorMessages.Any())
            {
                var errorMsg = string.Join("\n", errorMessages.Select(kv => $"ID: {kv.Key} - Errors: {string.Join(", ", kv.Value.Distinct())}"));
                var previewErrorMsg = errorMsg.Length > 500 ? errorMsg.Substring(0, 500) + "..." : errorMsg; // Limit preview to 500 characters

                var taskDialog = new TaskDialog("Invalid Pipes/Ducts")
                {
                    MainContent = "There are errors in placing some sleeves. Do you want to save these errors to a text file?",
                    ExpandedContent = previewErrorMsg,
                    CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
                };

                if (taskDialog.Show() == TaskDialogResult.Yes)
                {
                    // Prompt user to select file save location
                    var saveFileDialog = new SaveFileDialog
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
            else
            {
                TaskDialog.Show("Intersections", $"Placed {intersectionData.Count} スリーブ_SK instances at intersections.");
            }

            return Result.Succeeded;
        }

        private double GetBeamHeight(XYZ point1, XYZ point2, List<Element> structuralFramings)
        {
            foreach (var framing in structuralFramings)
            {
                var framingGeometry = framing.get_Geometry(new Options());
                if (framingGeometry == null) continue;

                var solids = GetSolidsFromGeometry(framingGeometry);

                foreach (var solid in solids)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face.Project(point1) != null && face.Project(point2) != null)
                        {
                            var boundingBox = solid.GetBoundingBox();
                            return boundingBox.Max.Z - boundingBox.Min.Z; // Assuming Z direction is the height
                        }
                    }
                }
            }

            return 0;
        }

        private List<Solid> GetSolidsFromGeometry(GeometryElement geometryElement)
        {
            var solids = new List<Solid>();

            foreach (var geomObj in geometryElement)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                {
                    solids.Add(solid);
                }
                else if (geomObj is GeometryInstance geomInstance)
                {
                    var instanceGeometry = geomInstance.GetInstanceGeometry();
                    solids.AddRange(GetSolidsFromGeometry(instanceGeometry));
                }
            }

            return solids;
        }

        private DirectShape CreateDirectShapeForBeamFace(Document doc, Solid solid, Face face)
        {
            // Get the bounding box of the face
            var boundingBox = face.GetBoundingBox();
            var min = boundingBox.Min;
            var max = boundingBox.Max;

            // Calculate the beam height (assuming Z direction is the height)
            var solidBoundingBox = solid.GetBoundingBox();
            var beamHeight = solidBoundingBox.Max.Z - solidBoundingBox.Min.Z;

            // Calculate the new dimensions for the direct shape
            var heightMargin = beamHeight / 4;
            var widthMargin = beamHeight;

            var adjustedMin = new UV(min.U + widthMargin, min.V + heightMargin);
            var adjustedMax = new UV(max.U - widthMargin, max.V - heightMargin);

            // Create a loop for the direct shape profile
            var profile = new List<Curve>
            {
                Line.CreateBound(face.Evaluate(adjustedMin), face.Evaluate(new UV(adjustedMin.U, adjustedMax.V))),
                Line.CreateBound(face.Evaluate(new UV(adjustedMin.U, adjustedMax.V)), face.Evaluate(adjustedMax)),
                Line.CreateBound(face.Evaluate(adjustedMax), face.Evaluate(new UV(adjustedMax.U, adjustedMin.V))),
                Line.CreateBound(face.Evaluate(new UV(adjustedMax.U, adjustedMin.V)), face.Evaluate(adjustedMin))
            };

            // Create a solid from the profile
            var curveLoop = CurveLoop.Create(profile);
            var directShapeSolid = GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { curveLoop }, face.ComputeNormal(UV.Zero), 10.0 / 304.8); // 10mm in feet

            // Create a direct shape
            var directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
            directShape.SetShape(new GeometryObject[] { directShapeSolid });
            directShape.SetName("Beam Intersection Zone");

            return directShape;
        }

        private bool IsPointWithinDirectShape(XYZ point, DirectShape directShape)
        {
            var boundingBox = directShape.get_BoundingBox(null);
            if (boundingBox == null) return false;

            return (point.X >= boundingBox.Min.X && point.X <= boundingBox.Max.X) &&
                   (point.Y >= boundingBox.Min.Y && point.Y <= boundingBox.Max.Y) &&
                   (point.Z >= boundingBox.Min.Z && point.Z <= boundingBox.Max.Z);
        }
    }
}
