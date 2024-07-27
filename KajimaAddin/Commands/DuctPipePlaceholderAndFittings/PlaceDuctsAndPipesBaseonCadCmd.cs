using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SKToolsAddins.Utils;
using Arc = Autodesk.Revit.DB.Arc;
using Curve = Autodesk.Revit.DB.Curve;
using Line = Autodesk.Revit.DB.Line;
using Solid = Autodesk.Revit.DB.Solid;
using Transaction = Autodesk.Revit.DB.Transaction;
using UnitUtils = SKToolsAddins.Utils.UnitUtils;
using MEPCurveUtils = SKToolsAddins.Utils.MEPCurveUtils;
using TextBox = System.Windows.Forms.TextBox;
using FlowDirection = System.Windows.Forms.FlowDirection;
using Form = System.Windows.Forms.Form;
using Panel = System.Windows.Forms.Panel;
using System.Windows.Controls;
using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;

namespace SKToolsAddins.Commands.DuctPipePlaceholderAndFittings
{
    [Transaction(TransactionMode.Manual)]
    public class PlaceDuctsAndPipesBaseonCadCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get the CAD link selected by the user
            var refLinkCad = uidoc.Selection.PickObject(ObjectType.Element, new ImportInstanceSelectionFilter(), "Select Link File");
            var selectedCadLink = doc.GetElement(refLinkCad) as ImportInstance;

            // Get only CAD lines and polylines in the CAD document
            var curveLayerMappings = GetCurveLayerMappings(selectedCadLink);

            Level level = uidoc.ActiveView.GenLevel;
            var customCurves = new List<CustomCurve>();

            // Placeholder types and system mappings based on CAD layer names
            var ductMappings = GetDuctMappings();
            var pipeMappings = GetPipeMappings();

            var systemTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(MechanicalSystemType))
                .Cast<MechanicalSystemType>()
                .ToDictionary(e => e.Name);

            var pipingSystemTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(PipingSystemType))
                .Cast<PipingSystemType>()
                .ToDictionary(e => e.Name);

            var ductTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(DuctType))
                .Cast<DuctType>()
                .ToDictionary(e => e.Name);

            var pipeTypes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .OfClass(typeof(PipeType))
                .Cast<PipeType>()
                .ToDictionary(e => e.Name);

            // Show input dialog to get user-defined offsets for relevant systems
            var offsets = GetSystemOffsets(curveLayerMappings, ductMappings, pipeMappings);

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Create duct and pipe placeholders from CAD lines and polylines");

                CreateMEPCurves(doc, level, curveLayerMappings, ductMappings, pipeMappings, ductTypes, pipeTypes, systemTypes, pipingSystemTypes, customCurves);

                var arcInBlock = CadUtils.GetArcsFromImportInstance(selectedCadLink);
                var customSolids = CreateSolidsFromArcs(doc, arcInBlock, level);
                var xPointsInsideBlock = ProcessCustomSolids(doc, level,customSolids, customCurves);

                var allCurves = customCurves.Select(c => c.MepCurve).ToList();
                var xPoints = MEPCurveUtils.FindIntersectionPoints(allCurves).Where(point => !xPointsInsideBlock.Any(blockPoint => blockPoint.IsAlmostEqualTo(point))).ToList();

                ProcessCurvesForFittings(doc, level, customCurves, xPoints, offsets);

                trans.Commit();
            }

            TaskDialog.Show("Success", "Successfully created duct and pipe from imported CAD lines and polylines.");
            return Result.Succeeded;
        }

        private void ProcessCurvesForFittings(Document doc, Level level, List<CustomCurve> customCurves, List<XYZ> xPoints, Dictionary<string, double> offsets)
        {
            foreach (var customCurve in customCurves.ToList())
            {
                var splitCurves = MEPCurveUtils.SplitCurve(doc, customCurve.MepCurve, xPoints, level);
                customCurve.SplitCurves.AddRange(splitCurves);

                foreach (var splitCurve in splitCurves)
                {
                    SetOffsetsAndSizeForSplitCurve(splitCurve, offsets);

                    var connectors = splitCurve.ConnectorManager.Connectors.Cast<Connector>().ToList();
                    foreach (var connector in connectors)
                    {
                        var xPoint = connector.Origin;
                        customCurve.XPointsConnectors.Add((xPoint, connector));
                    }
                }
            }

            var groupedPoints = GroupConnectorsByPoints(customCurves);

            foreach (var elbowPoint in groupedPoints.elbowPoints)
            {
                CreateElbowFitting(doc, elbowPoint.Connectors[0], elbowPoint.Connectors[1]);
            }

            foreach (var teePoint in groupedPoints.teePoints)
            {
                CreateTeeFitting(doc, teePoint.Connectors);
            }
        }

        private void SetOffsetsAndSizeForSplitCurve(MEPCurve splitCurve, Dictionary<string, double> offsets)
        {
            switch (splitCurve)
            {
                case Duct duct:
                    var ductSystemType = duct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString();
                    if (ductSystemType != null && offsets.TryGetValue(ductSystemType, out double ductOffset))
                    {
                        duct.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(UnitUtils.MmToFeet(ductOffset));
                        duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).Set(75 / 304.8);
                    }
                    break;
                case Pipe pipe:
                    var pipeSystemType = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                    if (pipeSystemType != null && offsets.TryGetValue(pipeSystemType, out double pipeOffset))
                    {
                        pipe.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(UnitUtils.MmToFeet(pipeOffset));
                        pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(75 / 304.8);
                    }
                    break;
            }
        }

        private void CreateMEPCurves(Document doc, Level level, Dictionary<Curve, string> curveLayerMappings, Dictionary<string, (string type, string system)> ductMappings, Dictionary<string, (string type, string system)> pipeMappings, Dictionary<string, DuctType> ductTypes, Dictionary<string, PipeType> pipeTypes, Dictionary<string, MechanicalSystemType> systemTypes, Dictionary<string, PipingSystemType> pipingSystemTypes, List<CustomCurve> customCurves)
        {
            foreach (var curve in curveLayerMappings.Keys)
            {
                if (!curveLayerMappings.TryGetValue(curve, out var layerName)) continue;

                if (ductMappings.TryGetValue(layerName, out var ductInfo))
                {
                    CreateDuct(doc, level, curve, ductInfo, ductTypes, systemTypes, customCurves);
                }
                else if (pipeMappings.TryGetValue(layerName, out var pipeInfo))
                {
                    CreatePipe(doc, level, curve, pipeInfo, pipeTypes, pipingSystemTypes, customCurves);
                }
            }
        }

        private void CreateDuct(Document doc, Level level, Curve curve, (string type, string system) ductInfo, Dictionary<string, DuctType> ductTypes, Dictionary<string, MechanicalSystemType> systemTypes, List<CustomCurve> customCurves)
        {
            var (ductTypeName, systemTypeName) = ductInfo;

            if (!ductTypes.TryGetValue(ductTypeName, out var ductType) || !systemTypes.TryGetValue(systemTypeName, out var systemType))
                return;

            XYZ startPoint = curve.GetEndPoint(0);
            XYZ endPoint = curve.GetEndPoint(1);

            if (startPoint.DistanceTo(endPoint) < 0.01)
                return;

            var duct = Duct.Create(doc, systemType.Id, ductType.Id, level.Id, startPoint, endPoint);
            customCurves.Add(new CustomCurve(duct));
        }

        private void CreatePipe(Document doc, Level level, Curve curve, (string type, string system) pipeInfo, Dictionary<string, PipeType> pipeTypes, Dictionary<string, PipingSystemType> pipingSystemTypes, List<CustomCurve> customCurves)
        {
            var (pipeTypeName, systemTypeName) = pipeInfo;

            if (!pipeTypes.TryGetValue(pipeTypeName, out var pipeType) || !pipingSystemTypes.TryGetValue(systemTypeName, out var systemType))
                return;

            XYZ startPoint = curve.GetEndPoint(0);
            XYZ endPoint = curve.GetEndPoint(1);

            if (startPoint.DistanceTo(endPoint) < 0.01)
                return;

            var pipe = Pipe.Create(doc, systemType.Id, pipeType.Id, level.Id, startPoint, endPoint);
            customCurves.Add(new CustomCurve(pipe));
        }

        private Dictionary<Curve, string> GetCurveLayerMappings(ImportInstance selectedCadLink)
        {
            var curveLayerMappings = new Dictionary<Curve, string>();
            var option = new Options();
            var geoElement = selectedCadLink.get_Geometry(option);
            var minCurveLength = 0.0016; // Minimum curve length in feet (1/16 inch)

            foreach (GeometryObject geoObject in geoElement)
            {
                if (geoObject is GeometryInstance geoInstance)
                {
                    var geoElement2 = geoInstance.GetInstanceGeometry();
                    foreach (GeometryObject geoObject2 in geoElement2)
                    {
                        if (geoObject2 is Line line)
                        {
                            if (line.Length >= minCurveLength)
                            {
                                curveLayerMappings[line] = CadUtils.GetLayerNameFromCurveOrPolyline(line, selectedCadLink);
                            }
                        }
                        else if (geoObject2 is PolyLine polyLine)
                        {
                            var points = polyLine.GetCoordinates();
                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                if (points[i].DistanceTo(points[i + 1]) >= minCurveLength)
                                {
                                    var segment = Line.CreateBound(points[i], points[i + 1]);
                                    curveLayerMappings[segment] = CadUtils.GetLayerNameFromCurveOrPolyline(polyLine, selectedCadLink);
                                }
                            }
                        }
                    }
                }
            }

            return curveLayerMappings;
        }

        private Dictionary<string, (string type, string system)> GetDuctMappings()
        {
            return new Dictionary<string, (string type, string system)>
            {
                { "M6SA", ("2D_00_丸ティー", "M06_給気_SA") },
                { "M6RA", ("2D_00_丸ティー", "M06_還気_RA") },
                { "M6EA", ("2D_00_丸ティー", "M06_排気_EA") },
                { "M6OA", ("2D_00_丸ティー", "M06_外気_OA") },
                { "M6PASS", ("2D_00_丸ティー", "M06_パス_PA") },
                { "M6SOA", ("2D_00_丸ティー", "M06_外気(処理外気)_SOA") },
                { "M6KEA", ("2D_00_丸ティー", "M06_厨房排気_KEA") }
            };
        }

        private Dictionary<string, (string type, string system)> GetPipeMappings()
        {
            return new Dictionary<string, (string type, string system)>
            {
                { "M5D", ("2D_00_排水_ドレン(空調)", "M05_ドレン(空調)_D") },
                { "M3R", ("2D_00_排水_冷媒_R", "M03_冷媒_R") },
                { "P1WATER", ("2D_00_加湿給水/C/CH", "M05_加湿給水") },
                { "M4C", ("2D_00_加湿給水/C/CH", "M04_冷水(往)_C") },
                { "M4CH", ("2D_00_加湿給水/C/CH", "M04_冷温水(往)_CH") }
            };
        }

        private Dictionary<string, double> GetSystemOffsets(Dictionary<Curve, string> curveLayerMappings, Dictionary<string, (string type, string system)> ductMappings, Dictionary<string, (string type, string system)> pipeMappings)
        {
            var offsets = new Dictionary<string, double>();
            var relevantSystems = new HashSet<string>();

            foreach (var layer in curveLayerMappings.Values)
            {
                if (ductMappings.TryGetValue(layer, out var ductInfo))
                {
                    relevantSystems.Add(ductInfo.system);
                }
                else if (pipeMappings.TryGetValue(layer, out var pipeInfo))
                {
                    relevantSystems.Add(pipeInfo.system);
                }
            }

            using (var form = new Form())
            {
                form.Text = "Input Offsets for Systems";
                form.Width = 400;
                form.Height = 400;
                form.StartPosition = FormStartPosition.CenterScreen;

                var panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };

                var tableLayoutPanel = new TableLayoutPanel
                {
                    ColumnCount = 2,
                    RowCount = relevantSystems.Count + 1,
                    Dock = DockStyle.Top,
                    AutoSize = true
                };

                tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
                tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));

                tableLayoutPanel.Controls.Add(new Label { Text = "System", TextAlign = System.Drawing.ContentAlignment.MiddleCenter }, 0, 0);
                tableLayoutPanel.Controls.Add(new Label { Text = "Offset (mm)", TextAlign = System.Drawing.ContentAlignment.MiddleCenter }, 1, 0);

                int rowIndex = 1;
                foreach (var system in relevantSystems)
                {
                    tableLayoutPanel.Controls.Add(new Label { Text = system, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, rowIndex);
                    var textBox = new TextBox { Tag = system, Text = "2800" };
                    tableLayoutPanel.Controls.Add(textBox, 1, rowIndex);
                    rowIndex++;
                }

                panel.Controls.Add(tableLayoutPanel);

                var buttonOk = new Button { Text = "OK", DialogResult = DialogResult.OK };
                var buttonCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };

                var flowLayoutPanel = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.RightToLeft,
                    Dock = DockStyle.Bottom,
                    AutoSize = true
                };

                flowLayoutPanel.Controls.Add(buttonCancel);
                flowLayoutPanel.Controls.Add(buttonOk);

                form.Controls.Add(panel);
                form.Controls.Add(flowLayoutPanel);

                if (form.ShowDialog() == DialogResult.OK)
                {
                    for (int i = 1; i < rowIndex; i++)
                    {
                        var textBox = tableLayoutPanel.GetControlFromPosition(1, i) as TextBox;
                        if (double.TryParse(textBox.Text, out double offset))
                        {
                            offsets[textBox.Tag.ToString()] = offset;
                        }
                    }
                }
            }

            return offsets;
        }

        public List<CustomSolid> CreateSolidsFromArcs(Document doc, IEnumerable<Arc> arcs, Level level)
        {
            var customSolids = new List<CustomSolid>();

            foreach (var arc in arcs)
            {
                double radius = arc.Radius + 30 / 304.8;
                XYZ centerPoint = arc.Center;
                //XYZ newCenterPoint = new XYZ(centerPoint.X, centerPoint.Y, centerPoint.Z - 2800 / 304.8);
                //double dist = 10000 / 304.8;
                if (level == null) continue;
                Solid arcBlockCheckIntersection = centerPoint.CreateCylinderUpAndDnByLevel(doc, radius, 1, level);
                //arcBlockCheckIntersection.BakeSolidToDirectShape(doc);
                var customSolid = new CustomSolid(arcBlockCheckIntersection);
                customSolids.Add(customSolid);
            }

            return customSolids;
        }

        private List<XYZ> ProcessCustomSolids(Document doc, Level level, List<CustomSolid> customSolids, List<CustomCurve> customCurves)
        {
            var xPointsInsideBlock = new List<XYZ>();

            foreach (var customSolid in customSolids)
            {
                foreach (var customCurve in customCurves.ToList())
                {
                    var curve = (customCurve.MepCurve.Location as LocationCurve)?.Curve;
                    if (curve == null)
                        continue;

                    var bottomZ = customSolid.Solid.GetBottomPlanarFace().Origin.Z;
                    var curveTransformed = curve.CreateTransformed(Transform.CreateTranslation(new XYZ(0, 0, bottomZ)));
                    var intersectionResult = customSolid.Solid.IntersectWithCurve(curveTransformed, new SolidCurveIntersectionOptions());

                    if (intersectionResult != null && intersectionResult.SegmentCount > 0)
                    {
                        customSolid.MepCurves.Add(customCurve.MepCurve);
                    }
                }

                var collinearCurves = MEPCurveUtils.GetCollinearCurves(customSolid.MepCurves);
                if (collinearCurves.Count == 0) continue;

                var mepCurvesOfCustomSolid = customSolid.MepCurves;
                try
                {
                    foreach (var (curve1, curve2) in collinearCurves)
                    {
                        if (!MEPCurveUtils.IsElementValid(curve1) || !MEPCurveUtils.IsElementValid(curve2))
                        {
                            continue;
                        }

                        var newCurve = MEPCurveUtils.CreateNewCurveFromCurvesCollinear(doc, curve1, curve2, level);
                        if (newCurve == null)
                        {
                            continue;
                        }

                        var collinearCurveIds = collinearCurves.SelectMany(cc => new[] { cc.Item1.Id, cc.Item2.Id }).Distinct().ToList();
                        customSolid.MepCurves = customSolid.MepCurves.Where(mc => collinearCurveIds.Contains(mc.Id)).ToList();

                        customCurves.Add(new CustomCurve(newCurve));
                        mepCurvesOfCustomSolid.Add(newCurve);

                        mepCurvesOfCustomSolid.RemoveAll(c => c.Id == curve1.Id || c.Id == curve2.Id);
                        customSolid.MepCurves.RemoveAll(c => c.Id == curve1.Id || c.Id == curve2.Id);
                        customCurves.RemoveAll(c => c.MepCurve.Id == curve1.Id || c.MepCurve.Id == curve2.Id);
                        doc.Delete(curve1.Id);
                        doc.Delete(curve2.Id);
                    }

                }
                catch 
                {
                }
                xPointsInsideBlock.AddRange(MEPCurveUtils.FindIntersectionPoints(mepCurvesOfCustomSolid));
            }

            return xPointsInsideBlock;
        }

        private (List<(XYZ Point, List<Connector> Connectors)> teePoints, List<(XYZ Point, List<Connector> Connectors)> elbowPoints) GroupConnectorsByPoints(List<CustomCurve> customCurves)
        {
            var allXPointsConnectors = customCurves.SelectMany(c => c.XPointsConnectors);
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
                    Connectors = g.Select(p => p.Connector).ToList()
                })
                .ToList();

            var teePoints = new List<(XYZ Point, List<Connector> Connectors)>();
            var elbowPoints = new List<(XYZ Point, List<Connector> Connectors)>();

            foreach (var group in groupedPoints)
            {
                if (group.Connectors.Count == 3)
                {
                    teePoints.Add((new XYZ(group.Point.X, group.Point.Y, group.Point.Z), group.Connectors));
                }
                else if (group.Connectors.Count == 2)
                {
                    elbowPoints.Add((new XYZ(group.Point.X, group.Point.Y, group.Point.Z), group.Connectors));
                }
            }

            return (teePoints, elbowPoints);
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
            try
            {
                if (connectors.Count == 3)
                {
                    var conn1 = connectors[0];
                    var conn2 = connectors[1];
                    var conn3 = connectors[2];

                    var curve1 = conn1.Owner as MEPCurve;
                    var curve2 = conn2.Owner as MEPCurve;
                    var curve3 = conn3.Owner as MEPCurve;

                    if (MEPCurveUtils.AreCurvesCollinear(curve1, curve2))
                    {
                        doc.Create.NewTeeFitting(conn1, conn2, conn3);
                    }
                    else if (MEPCurveUtils.AreCurvesCollinear(curve1, curve3))
                    {
                        doc.Create.NewTeeFitting(conn1, conn3, conn2);
                    }
                    else if (MEPCurveUtils.AreCurvesCollinear(curve2, curve3))
                    {
                        doc.Create.NewTeeFitting(conn2, conn3, conn1);
                    }
                }
            }
            catch
            {
                // Skip if the tee fitting creation fails
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
                MepCurve = mepCurve;
                XPoints = new List<XYZ>();
                SplitCurves = new List<MEPCurve>();
                XPointsConnectors = new List<(XYZ XPoints, Connector Connector)>();
            }
        }

        public class CustomSolid
        {
            public Solid Solid { get; set; }
            public List<MEPCurve> MepCurves { get; set; }

            public CustomSolid(Solid solid)
            {
                Solid = solid;
                MepCurves = new List<MEPCurve>();
            }
        }

        class ImportInstanceSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is ImportInstance;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
    }
}
