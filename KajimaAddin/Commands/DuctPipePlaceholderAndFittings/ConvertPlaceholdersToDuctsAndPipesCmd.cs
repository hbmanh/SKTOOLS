using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using SKToolsAddins.Utils;

namespace SKToolsAddins.Commands.DuctPipePlaceholderAndFittings
{
    [Transaction(TransactionMode.Manual)]
    public class ConvertPlaceholdersToDuctsAndPipesCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            Level level = uidoc.ActiveView.GenLevel;

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Convert Placeholders to Pipes and Ducts");

                List<MEPCurve> mepCurves = new List<MEPCurve>();

                // Thu thập tất cả các Pipe Placeholder và Duct Placeholder từ level hiện tại

                List<Pipe> pipePlaceholders = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PlaceHolderPipes)
                    .OfClass(typeof(Pipe))
                    .WhereElementIsNotElementType()
                    .Cast<Pipe>()
                    .Where(d => d.ReferenceLevel.Id == level.Id)
                    .ToList();

                List<Duct> ductPlaceholders = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PlaceHolderDucts)
                    .OfClass(typeof(Duct))
                    .WhereElementIsNotElementType()
                    .Cast<Duct>()
                    .Where(d => d.ReferenceLevel.Id == level.Id)
                    .ToList();
                mepCurves.AddRange(pipePlaceholders);
                mepCurves.AddRange(ductPlaceholders);

        var xPoints = MEPCurveUtils.FindIntersectionPoints(mepCurves).ToList();
                List<CustomCurve> customCurves = new List<CustomCurve>();
                foreach (var mepCurve in mepCurves)
                {
                    var customCurve = new CustomCurve(mepCurve);
                    // Chia các MEPCurve dựa trên xPoints
                    var splitCurves = MEPCurveUtils.SplitCurve(doc, mepCurve, xPoints, level);
                    customCurve.SplitCurves.AddRange(splitCurves);

                    // Add connector vào Xpoints
                    foreach (var splitCurve in splitCurves)
                    {
                        var connectors = splitCurve.ConnectorManager.Connectors.Cast<Connector>().ToList();
                        foreach (var connector in connectors)
                        {
                            var xPoint = connector.Origin;
                            customCurve.XPoints.Add(xPoint);
                            customCurve.XPointsConnectors.Add((xPoint, connector));
                        }
                    }
                    customCurves.Add(customCurve);
                }
                var allXPointsConnectors = customCurves.SelectMany(c => c.XPointsConnectors);
                // Nhóm các XPointsConnectors theo tọa độ X, Y, Z sau khi làm tròn
                var groupedPoints = allXPointsConnectors
                    .GroupBy(p => new
                    {
                        X = Math.Round(p.XPoints.X, 0),
                        Y = Math.Round(p.XPoints.Y, 0),
                        Z = Math.Round(p.XPoints.Z, 0)
                    })
                    .Select(g => new
                    {
                        Point = g.Key,
                        // Danh sách các Connectors tương ứng với điểm đó
                        Connectors = g.Select(p => p.Connector).ToList()
                    })
                    .ToList();
                var isTeePoints = new List<(XYZ Point, List<Connector> Connectors)>();
                var isElbowPoints = new List<(XYZ Point, List<Connector> Connectors)>();
                // Tìm kiếm các điểm Tee và Elbow
                foreach (var group in groupedPoints)
                {
                    if (group.Connectors.Count == 3)
                    {
                        isTeePoints.Add((new XYZ(group.Point.X, group.Point.Y, group.Point.Z), group.Connectors));
                    }
                    else if (group.Connectors.Count == 2)
                    {
                        isElbowPoints.Add((new XYZ(group.Point.X, group.Point.Y, group.Point.Z), group.Connectors));
                    }
                }

                // Đặt Elbow connector tại các điểm elbow
                foreach (var elbowPoint in isElbowPoints)
                {
                    CreateElbowFitting(doc, elbowPoint.Connectors[0], elbowPoint.Connectors[1]);
                }

                // Đặt Tee connector tại các điểm T
                foreach (var teePoint in isTeePoints)
                {
                    CreateTeeFitting(doc, teePoint.Connectors);
                }

                trans.Commit();
            }

            TaskDialog.Show("Success", "Successfully converted Placeholders to Pipes and Ducts");
            return Result.Succeeded;
        }
        private void CreateElbowFitting(Document doc, Connector connector1, Connector connector2)
        {
            try
            {
                doc.Create.NewElbowFitting(connector1, connector2);
            }
            catch
            {
                // Skip if the elbow fitting already exists
            }
        }

        private void CreateTeeFitting(Document doc, List<Connector> connectors)
        {
            if (connectors.Count == 3)
            {
                // Xác định ba connector
                Connector conn1 = connectors[0];
                Connector conn2 = connectors[1];
                Connector conn3 = connectors[2];

                // Lấy MEPCurve tương ứng của các connector
                MEPCurve curve1 = conn1.Owner as MEPCurve;
                MEPCurve curve2 = conn2.Owner as MEPCurve;
                MEPCurve curve3 = conn3.Owner as MEPCurve;

                // Kiểm tra xem hai connector nào nằm trên cùng một đường thẳng
                try
                {
                    if (MEPCurveUtils.AreCurvesCollinear(curve1, curve2))
                    {
                        // conn3 là branch
                        doc.Create.NewTeeFitting(conn1, conn2, conn3);
                    }
                    else if (MEPCurveUtils.AreCurvesCollinear(curve1, curve3))
                    {
                        // conn2 là branch
                        doc.Create.NewTeeFitting(conn1, conn3, conn2);
                    }
                    else if (MEPCurveUtils.AreCurvesCollinear(curve2, curve3))
                    {
                        // conn1 là branch
                        doc.Create.NewTeeFitting(conn2, conn3, conn1);
                    }
                }
                catch
                {
                    // Skip if the tee fitting creation fails
                }
            }
        }

        class CustomCurve
        {
            public MEPCurve MepCurve { get; set; }
            public List<XYZ> XPoints { get; set; }
            public List<MEPCurve> SplitCurves { get; set; }
            public List<(XYZ XPoints, Connector Connector)> XPointsConnectors { get; set; }

            public CustomCurve(MEPCurve mepCurve)
            {
                this.MepCurve = mepCurve;
                XPoints = new List<XYZ>();
                SplitCurves = new List<MEPCurve>();
                XPointsConnectors = new List<(XYZ XPoints, Connector Connector)>();
            }
        }
    }

}
