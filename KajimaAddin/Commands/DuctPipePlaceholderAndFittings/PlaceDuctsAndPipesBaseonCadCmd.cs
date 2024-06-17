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
            Dictionary<Curve, string> curveLayerMappings = new Dictionary<Curve, string>();

            // Defaults to medium detail, no references and no view.
            Options option = new Options();
            GeometryElement geoElement = selectedCadLink.get_Geometry(option);

            foreach (GeometryObject geoObject in geoElement)
            {
                if (!(geoObject is GeometryInstance geoInstance)) continue;
                GeometryElement geoElement2 = geoInstance.GetInstanceGeometry();
                foreach (GeometryObject geoObject2 in geoElement2)
                {
                    if (geoObject2 is Line line)
                    {
                        curveLayerMappings[line] = CadUtils.GetLayerNameFromCurveOrPolyline(line, selectedCadLink);
                    }
                    else if (geoObject2 is PolyLine polyLine)
                    {
                        IList<XYZ> points = polyLine.GetCoordinates();
                        for (int i = 0; i < points.Count - 1; i++)
                        {
                            Line segment = Line.CreateBound(points[i], points[i + 1]);
                            curveLayerMappings[segment] = CadUtils.GetLayerNameFromCurveOrPolyline(polyLine, selectedCadLink); // Get the layer name from the polyline
                        }
                    }
                }
            }

            Level level = uidoc.ActiveView.GenLevel;
            List<CustomCurve> customCurves = new List<CustomCurve>();

            // Placeholder types and system mappings based on CAD layer names
            var ductMappings = new Dictionary<string, (string type, string system)>
            {
                { "M6SA", ("2D_00_丸ティー", "M06_給気_SA") },
                { "M6RA", ("2D_00_丸ティー", "M06_還気_RA") },
                { "M6EA", ("2D_00_丸ティー", "M06_排気_EA") },
                { "M6OA", ("2D_00_丸ティー", "M06_外気_OA") },
                { "M6PASS", ("2D_00_丸ティー", "M06_パス_PA") },
                { "M6SOA", ("2D_00_丸ティー", "M06_外気(処理外気)_SOA") },
                { "M6KEA", ("2D_00_丸ティー", "M06_厨房排気_KEA") }
            };

            var pipeMappings = new Dictionary<string, (string type, string system)>
            {
                { "M5D", ("2D_00_排水_ドレン(空調)", "M05_ドレン(空調)_D") },
                { "M3R", ("2D_00_排水_冷媒_R", "M03_冷媒_R") },
                { "P1WATER", ("2D_00_加湿給水/C/CH", "M05_加湿給水") },
                { "M4C", ("2D_00_加湿給水/C/CH", "M04_冷水(往)_C") },
                { "M4CH", ("2D_00_加湿給水/C/CH", "M04_冷温水(往)_CH") }
            };

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

                foreach (var curve in curveLayerMappings.Keys)
                {
                    if (!curveLayerMappings.TryGetValue(curve, out var layerName))
                        continue;

                    if (ductMappings.TryGetValue(layerName, out var ductInfo))
                    {
                        var (ductTypeName, systemTypeName) = ductInfo;

                        if (!ductTypes.TryGetValue(ductTypeName, out var ductType) || !systemTypes.TryGetValue(systemTypeName, out var systemType))
                            continue;

                        XYZ startPoint = curve.GetEndPoint(0);
                        XYZ endPoint = curve.GetEndPoint(1);

                        if (startPoint.DistanceTo(endPoint) < 0.01)
                            continue;

                        var duct = Duct.Create(doc, systemType.Id, ductType.Id, level.Id, startPoint, endPoint);
                        customCurves.Add(new CustomCurve(duct));
                    }
                    else if (pipeMappings.TryGetValue(layerName, out var pipeInfo))
                    {
                        var (pipeTypeName, systemTypeName) = pipeInfo;

                        if (!pipeTypes.TryGetValue(pipeTypeName, out var pipeType) || !pipingSystemTypes.TryGetValue(systemTypeName, out var systemType))
                            continue;

                        XYZ startPoint = curve.GetEndPoint(0);
                        XYZ endPoint = curve.GetEndPoint(1);

                        if (startPoint.DistanceTo(endPoint) < 0.01)
                            continue;

                        var pipe = Pipe.Create(doc, systemType.Id, pipeType.Id, level.Id, startPoint, endPoint);
                        customCurves.Add(new CustomCurve(pipe));
                    }
                }

                var arcInBlock = CadUtils.GetArcsFromImportInstance(selectedCadLink);
                List<XYZ> xPointsInsideBlock = new List<XYZ>();
                var isTeePoints = new List<(XYZ Point, List<Connector> Connectors)>();
                var isElbowPoints = new List<(XYZ Point, List<Connector> Connectors)>();
                List<CustomSolid> customSolids = CreateSolidsFromArcs(doc, arcInBlock, level);
                // Create newcurve to connect via block, and get intersection of 2 curve after create new curve in block
                foreach (var customSolid in customSolids)
                {
                    foreach (var customCurve in customCurves.ToList())
                    {
                        Curve curve = (customCurve.MepCurve.Location as LocationCurve)?.Curve;
                        if (curve == null)
                            continue;
                        var arcBlockCheckIntersection = customSolid.Solid;
                        var intersectionResult = arcBlockCheckIntersection.IntersectWithCurve(curve, new SolidCurveIntersectionOptions());
                        if (intersectionResult != null && intersectionResult.SegmentCount > 0)
                        {
                            customSolid.MepCurves.Add(customCurve.MepCurve);
                        }
                        var collinearCurves = MEPCurveUtils.GetCollinearCurves(customSolid.MepCurves);
                        if (collinearCurves.Count == 0) continue;
                        List<MEPCurve> mepCurvesOfCustomSolid = customSolid.MepCurves;
                        foreach (var (curve1, curve2) in collinearCurves)
                        {
                            var newCurve = MEPCurveUtils.CreateNewCurveFromCurvesCollinear(doc, curve1, curve2, level);
                            // Loại bỏ các MEPCurve không thẳng hàng ra khỏi customSolid.MepCurves
                            var collinearCurveIds = collinearCurves.SelectMany(cc => new[] { cc.Item1.Id, cc.Item2.Id }).Distinct().ToList();
                            customSolid.MepCurves = customSolid.MepCurves.Where(mc => collinearCurveIds.Contains(mc.Id)).ToList();
                            if (newCurve == null) continue;
                            customCurves.Add(new CustomCurve(newCurve));
                            mepCurvesOfCustomSolid.Add(newCurve); // To check intersection point inside block
                            mepCurvesOfCustomSolid.RemoveAll(c => c.Id == curve1.Id || c.Id == curve2.Id);
                            customSolid.MepCurves.RemoveAll(c => c.Id == curve1.Id || c.Id == curve2.Id);
                            customCurves.RemoveAll(c => c.MepCurve.Id == curve1.Id || c.MepCurve.Id == curve2.Id);
                            doc.Delete(curve1.Id);
                            doc.Delete(curve2.Id);
                        }
                        xPointsInsideBlock.AddRange(MEPCurveUtils.FindIntersectionPoints(mepCurvesOfCustomSolid));
                    }
                }

                var allCurves = customCurves.Select(c => c.MepCurve).ToList();
                var xPoints = MEPCurveUtils.FindIntersectionPoints(allCurves).Where(point => !xPointsInsideBlock.Any(blockPoint => blockPoint.IsAlmostEqualTo(point))).ToList();

                foreach (var customCurve in customCurves.ToList())
                {
                    // Chia các MEPCurve dựa trên xPoints
                    var splitCurves = MEPCurveUtils.SplitCurve(doc, customCurve.MepCurve, xPoints, level);
                    customCurve.SplitCurves.AddRange(splitCurves);

                    // Add connector vào Xpoints
                    foreach (var splitCurve in splitCurves)
                    {
                        var connectors = splitCurve.ConnectorManager.Connectors.Cast<Connector>().ToList();
                        foreach (var connector in connectors)
                        {
                            var xPoint = connector.Origin;
                            customCurve.XPointsConnectors.Add((xPoint, connector));
                        }
                    }
                }
                // Group points to classify TeePoints and ElbowPoints

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
                foreach (var elbowPoint in isElbowPoints)
                {
                    CreateElbowFitting(doc, elbowPoint.Connectors[0], elbowPoint.Connectors[1]);
                }
                // Đặt Tee connector tại các điểm T
                foreach (var teePoint in isTeePoints)
                {
                    CreateTeeFitting(doc, teePoint.Connectors);
                }
               
                foreach (var customCurve in customCurves)
                {
                    var mepCurve = customCurve.MepCurve;

                    try
                    {
                        switch (mepCurve)
                        {
                            case Duct _:
                            {
                                var systemType = mepCurve.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsString();
                                if (offsets.TryGetValue(systemType, out double offset))
                                {
                                    mepCurve.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(UnitUtils.MmToFeet(offset));
                                }

                                break;
                            }
                            case Pipe _:
                            {
                                var systemType = mepCurve.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsString();
                                if (offsets.TryGetValue(systemType, out double offset))
                                {
                                    mepCurve.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(UnitUtils.MmToFeet(offset));
                                }

                                break;
                            }
                        }
                    }
                    catch
                    {
                        //
                    }
                }
                trans.Commit();
            }

            TaskDialog.Show("Success", "Successfully created duct and pipe from imported CAD lines and polylines.");
            return Result.Succeeded;
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

            using (Form form = new Form())
            {
                form.Text = "Input Offsets for Systems";
                form.Width = 400;
                form.Height = 400;
                form.StartPosition = FormStartPosition.CenterScreen;

                Panel panel = new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true
                };

                TableLayoutPanel tableLayoutPanel = new TableLayoutPanel
                {
                    ColumnCount = 2,
                    RowCount = relevantSystems.Count + 1,
                    Dock = DockStyle.Top,
                    AutoSize = true
                };

                // Adjust the column widths
                tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
                tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));

                // Add header row
                tableLayoutPanel.Controls.Add(new Label() { Text = "System", TextAlign = System.Drawing.ContentAlignment.MiddleCenter }, 0, 0);
                tableLayoutPanel.Controls.Add(new Label() { Text = "Offset (mm)", TextAlign = System.Drawing.ContentAlignment.MiddleCenter }, 1, 0);

                // Add rows for relevant systems
                int rowIndex = 1;
                foreach (var system in relevantSystems)
                {
                    tableLayoutPanel.Controls.Add(new Label() { Text = system, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, rowIndex);
                    TextBox textBox = new TextBox() { Tag = system, Text = "2800" }; // Set default value to 2800
                    tableLayoutPanel.Controls.Add(textBox, 1, rowIndex);
                    rowIndex++;
                }

                panel.Controls.Add(tableLayoutPanel);

                // Add OK and Cancel buttons
                Button buttonOk = new Button() { Text = "OK", DialogResult = DialogResult.OK };
                Button buttonCancel = new Button() { Text = "Cancel", DialogResult = DialogResult.Cancel };

                FlowLayoutPanel flowLayoutPanel = new FlowLayoutPanel()
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
            List<CustomSolid> customSolids = new List<CustomSolid>();

            foreach (Arc arc in arcs)
            {
                double radius = arc.Radius + 10 / 304.8;
                XYZ centerPoint = arc.Center;
                double additionalHeight = 3200 / 304.8; 
                XYZ newCenterPoint = new XYZ(centerPoint.X, centerPoint.Y, centerPoint.Z + additionalHeight);
                if (level == null) continue;
                Solid arcBlockCheckIntersection = newCenterPoint.CreateCylinderUpAndDnByLevel(doc, radius, 1, level);
                //arcBlockCheckIntersection.BakeSolidToDirectShape(doc);
                CustomSolid customSolid = new CustomSolid(arcBlockCheckIntersection);
                customSolids.Add(customSolid);
            }

            return customSolids;
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
            // Kiểm tra xem hai connector nào nằm trên cùng một đường thẳng
            try
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
                this.MepCurve = mepCurve;
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
