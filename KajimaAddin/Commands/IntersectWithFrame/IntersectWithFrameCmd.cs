﻿using Autodesk.Revit.Attributes;
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
                        if (framingGeometry == null)
                            continue;

                        List<Solid> solids = GetSolidsFromGeometry(framingGeometry);

                        foreach (Solid solid in solids)
                        {
                            foreach (Face face in solid.Faces)
                            {
                                IntersectionResultArray resultArray;
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
                FamilySymbol sleeveSymbol = new FilteredElementCollector(doc)
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
                    for (int i = 0; i < points.Count; i += 2)
                    {
                        if (i + 1 < points.Count)
                        {
                            // Calculate midpoint and direction
                            XYZ point1 = points[i];
                            XYZ point2 = points[i + 1];
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

                            // Set the parameter OD to the diameter of the pipe/duct + 50mm
                            Parameter diameterParam = pipeOrDuct.LookupParameter("Diameter");
                            if (diameterParam != null && diameterParam.HasValue)
                            {
                                double pipeDiameter = diameterParam.AsDouble(); // Revit stores length units in feet by default
                                double sleeveDiameter = pipeDiameter + UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Millimeters); // Adding 50mm and converting to feet
                                Parameter odParam = sleeveInstance.LookupParameter("OD");
                                if (odParam != null)
                                {
                                    odParam.Set(sleeveDiameter);
                                }
                            }
                        }
                    }
                }

                trans.Commit();
            }

            TaskDialog.Show("Intersections", $"Placed {intersectionData.Count} スリーブ_SK instances at intersections.");
            return Result.Succeeded;
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
    }
}
