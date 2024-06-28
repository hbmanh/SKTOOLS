using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
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
            var structuralFramings = new List<Element>();
            foreach (var linkedDoc in linkedDocs)
            {
                var framings = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .ToElements();
                structuralFramings.AddRange(framings);
            }

            // Dictionary to store intersection results with midpoint and direction
            var intersectionData = new Dictionary<ElementId, List<XYZ>>();

            // Place the スリーブ_SK family instances at the midpoint of intersection points
            using (Transaction trans = new Transaction(doc, "Place Sleeves"))
            {
                trans.Start();
                foreach (var pipeOrDuct in pipesAndDucts)
                {
                    var pipeOrDuctCurve = (pipeOrDuct.Location as LocationCurve)?.Curve;
                    if (pipeOrDuctCurve == null)
                        continue;

                    foreach (var framing in structuralFramings)
                    {
                        var framingGeometry = framing.get_Geometry(new Options());

                        // Extract solids from the geometry
                        List<Solid> framingSolids = new List<Solid>();
                        foreach (GeometryObject geomObj in framingGeometry)
                        {
                            if (geomObj is Solid solid && solid.Volume > 0)
                            {
                                framingSolids.Add(solid);
                            }
                            else if (geomObj is GeometryInstance geomInstance)
                            {
                                var instanceGeometry = geomInstance.GetInstanceGeometry();
                                foreach (var instanceGeomObj in instanceGeometry)
                                {
                                    if (instanceGeomObj is Solid instanceSolid && instanceSolid.Volume > 0)
                                    {
                                        framingSolids.Add(instanceSolid);
                                    }
                                }
                            }
                        }

                        // Check intersections
                        foreach (Solid framingSolid in framingSolids)
                        {
                            foreach (Face face in framingSolid.Faces)
                            {
                                var resultArray = new IntersectionResultArray();
                                if (face.Intersect(pipeOrDuctCurve, out resultArray) == SetComparisonResult.Overlap)
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

                // Load the スリーブ_SK family symbol
                FamilySymbol sleeveSymbol = null;
                var pipeAccessories = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeAccessory)
                    .OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(symbol => symbol.Family.Name == "スリーブ_SK");

                if (pipeAccessories != null)
                {
                    sleeveSymbol = pipeAccessories;
                }
                else
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
                    var points = entry.Value;
                    if (points.Count >= 2)
                    {
                        // Calculate midpoint and direction
                        XYZ point1 = points[0];
                        XYZ point2 = points[1];
                        XYZ midpoint = (point1 + point2) / 2;
                        XYZ direction = (point2 - point1).Normalize();

                        // Place the sleeve instance
                        FamilyInstance sleeveInstance = doc.Create.NewFamilyInstance(midpoint, sleeveSymbol, StructuralType.NonStructural);

                        // Rotate the sleeve to be parallel with the direction vector plus an additional 90 degrees
                        Line axis = Line.CreateBound(midpoint, midpoint + XYZ.BasisZ);
                        double angle = XYZ.BasisX.AngleTo(direction);
                        double additionalRotation = Math.PI / 2; // 90 degrees in radians
                        ElementTransformUtils.RotateElement(doc, sleeveInstance.Id, axis, angle + additionalRotation);

                        // Set the parameter L to the distance between the intersection points
                        Parameter lengthParam = sleeveInstance.LookupParameter("L");
                        if (lengthParam != null)
                        {
                            lengthParam.Set(point1.DistanceTo(point2));
                        }
                    }
                }

                trans.Commit();
            }

            TaskDialog.Show("Intersections", $"Placed {intersectionData.Count} スリーブ_SK instances at intersections.");
            return Result.Succeeded;
        }
    }
}
