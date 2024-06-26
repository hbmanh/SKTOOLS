//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.UI;
//using System.Collections.Generic;
//using System.Linq;

//namespace SKToolsAddins.Commands.IntersectWithFrame
//{
//    [Transaction(TransactionMode.Manual)]
//    public class IntersectWithFrameCmd : IExternalCommand
//    {
//        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//        {
//            UIApplication uiapp = commandData.Application;
//            UIDocument uidoc = uiapp.ActiveUIDocument;
//            Document doc = uidoc.Document;
//            View activeView = uidoc.ActiveView;

//            // Collect pipes and ducts in the active view
//            List<Element> pipes = GetElementsInView(doc, activeView.Id, BuiltInCategory.OST_PipeCurves);
//            List<Element> ducts = GetElementsInView(doc, activeView.Id, BuiltInCategory.OST_DuctCurves);

//            // Collect linked documents containing structural framing
//            List<RevitLinkInstance> linkedDocuments = GetLinkedDocumentsWithFraming(doc);

//            List<XYZ> intersectionPoints = new List<XYZ>();

//            // Process intersections for each linked document
//            foreach (RevitLinkInstance linkInstance in linkedDocuments)
//            {
//                Document linkDoc = linkInstance.GetLinkDocument();
//                if (linkDoc != null)
//                {
//                    Transform linkTransform = linkInstance.GetTransform();
//                    FilteredElementCollector framingCollector = new FilteredElementCollector(linkDoc, activeView.Id)
//                        .OfCategory(BuiltInCategory.OST_StructuralFraming)
//                        .OfClass(typeof(FamilyInstance));

//                    List<Element> structuralFramings = new List<Element>(framingCollector);

//                    // Check for intersections between pipes and structural framings
//                    intersectionPoints.AddRange(GetIntersections(pipes, structuralFramings, linkTransform));

//                    // Check for intersections between ducts and structural framings
//                    intersectionPoints.AddRange(GetIntersections(ducts, structuralFramings, linkTransform));
//                }
//            }

//            // Handle the intersection points as needed, e.g., store them in a parameter or highlight in UI
//            TaskDialog.Show("Intersections", $"Found {intersectionPoints.Count} intersections in the active view.");

//            return Result.Succeeded;
//        }

//        private List<Element> GetElementsInView(Document doc, ElementId viewId, BuiltInCategory category)
//        {
//            FilteredElementCollector collector = new FilteredElementCollector(doc, viewId)
//                .OfCategory(category)
//                .OfClass(typeof(MEPCurve));
//            return collector.ToList();
//        }

//        private List<RevitLinkInstance> GetLinkedDocumentsWithFraming(Document doc)
//        {
//            FilteredElementCollector linkCollector = new FilteredElementCollector(doc)
//                .OfClass(typeof(RevitLinkInstance));

//            List<RevitLinkInstance> linkedDocuments = new List<RevitLinkInstance>();
//            foreach (Element linkElement in linkCollector)
//            {
//                if (linkElement is RevitLinkInstance linkInstance)
//                {
//                    Document linkDoc = linkInstance.GetLinkDocument();
//                    if (linkDoc != null)
//                    {
//                        FilteredElementCollector framingCollector = new FilteredElementCollector(linkDoc)
//                            .OfCategory(BuiltInCategory.OST_StructuralFraming)
//                            .OfClass(typeof(FamilyInstance));

//                        if (framingCollector.GetElementCount() > 0)
//                        {
//                            linkedDocuments.Add(linkInstance);
//                        }
//                    }
//                }
//            }
//            return linkedDocuments;
//        }

//        private List<XYZ> GetIntersections(List<Element> mepElements, List<Element> structuralFramings, Transform linkTransform)
//        {
//            List<XYZ> intersectionPoints = new List<XYZ>();

//            foreach (Element mepElement in mepElements)
//            {
//                Curve mepCurve = GetElementCurve(mepElement);
//                if (mepCurve != null)
//                {
//                    foreach (Element framing in structuralFramings)
//                    {
//                        Solid framingSolid = GetElementSolid(framing);
//                        if (framingSolid != null)
//                        {
//                            Solid transformedSolid = SolidUtils.CreateTransformed(framingSolid, linkTransform);
//                            SolidCurveIntersection intersection = transformedSolid.IntersectWithCurve(mepCurve, new SolidCurveIntersectionOptions());
//                            if (intersection != null)
//                            {
//                                for (int i = 0; i < intersection.SegmentCount; i++)
//                                {
//                                    CurveSegment segment = intersection.GetCurveSegment(i);
//                                    IntersectionResultArray results = segment.ComputeCurveIntersections(mepCurve, false, false, 0.0);
//                                    foreach (IntersectionResult result in results)
//                                    {
//                                        intersectionPoints.Add(result.XYZPoint);
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }
//            }

//            return intersectionPoints;
//        }

//        private Curve GetElementCurve(Element element)
//        {
//            LocationCurve locCurve = element.Location as LocationCurve;
//            if (locCurve != null)
//            {
//                return locCurve.Curve;
//            }
//            return null;
//        }

//        private Solid GetElementSolid(Element element)
//        {
//            GeometryElement geomElement = element.get_Geometry(new Options());
//            foreach (GeometryObject geomObj in geomElement)
//            {
//                if (geomObj is GeometryInstance geomInstance)
//                {
//                    GeometryElement instanceGeom = geomInstance.GetInstanceGeometry();
//                    foreach (GeometryObject instanceObj in instanceGeom)
//                    {
//                        if (instanceObj is Solid solid && solid.Volume > 0)
//                        {
//                            return solid;
//                        }
//                    }
//                }
//                else if (geomObj is Solid solid && solid.Volume > 0)
//                {
//                    return solid;
//                }
//            }
//            return null;
//        }
//    }
//}
