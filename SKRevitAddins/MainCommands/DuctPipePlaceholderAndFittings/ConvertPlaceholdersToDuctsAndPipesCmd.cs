using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;
using Form = System.Windows.Forms.Form;
using Panel = System.Windows.Forms.Panel;
using TextBox = System.Windows.Forms.TextBox;
using UnitUtils = SKRevitAddins.Utils.UnitUtils;

namespace SKRevitAddins.DuctPipePlaceholderAndFittings
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

            // Lấy lựa chọn của người dùng
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            List<MEPCurve> mepCurves = new List<MEPCurve>();
            List<Element> selectedElements = new List<Element>();

            if (selectedIds.Count > 0)
            {
                // Nếu có lựa chọn
                selectedElements = selectedIds.Select(id => doc.GetElement(id)).ToList();
                mepCurves.AddRange(selectedElements.OfType<Pipe>().Where(d => d.ReferenceLevel.Id == level.Id));
                mepCurves.AddRange(selectedElements.OfType<Duct>().Where(d => d.ReferenceLevel.Id == level.Id));
            }
            else
            {
                // Nếu không có lựa chọn, lấy tất cả các đối tượng trong Active View
                var pipePlaceholders = new FilteredElementCollector(doc, uidoc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_PlaceHolderPipes)
                    .OfClass(typeof(Pipe))
                    .WhereElementIsNotElementType()
                    .Cast<Pipe>()
                    .Where(d => d.ReferenceLevel.Id == level.Id)
                    .ToList();

                var ductPlaceholders = new FilteredElementCollector(doc, uidoc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_PlaceHolderDucts)
                    .OfClass(typeof(Duct))
                    .WhereElementIsNotElementType()
                    .Cast<Duct>()
                    .Where(d => d.ReferenceLevel.Id == level.Id)
                    .ToList();

                mepCurves.AddRange(pipePlaceholders);
                mepCurves.AddRange(ductPlaceholders);
                selectedElements.AddRange(pipePlaceholders);
                selectedElements.AddRange(ductPlaceholders);
            }

            // Hiển thị hộp thoại nhập để lấy cao độ system cho các đối tượng được chọn hoặc trong active view
            var offsets = GetSystemOffsets(doc, selectedElements);

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Convert Placeholders to Pipes and Ducts");
                List<MEPCurve> newCurves = new List<MEPCurve>();
                foreach (var pipe in mepCurves.OfType<Pipe>())
                {
                    if (pipe == null) continue;
                    XYZ startPoint = (pipe.Location as LocationCurve)?.Curve.GetEndPoint(0);
                    XYZ endPoint = (pipe.Location as LocationCurve)?.Curve.GetEndPoint(1);
                    var systemId = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
                    var pipeSize = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsDouble(); // lấy kích thước
                    double offset = offsets.ContainsKey(systemId) ? offsets[systemId] : 2800; // Default offset if not provided
                    if (startPoint != null && endPoint != null)
                    {
                        var newPipe = Pipe.Create(doc, systemId, pipe.GetTypeId(), level.Id, startPoint, endPoint);
                        if (pipeSize.HasValue)
                        {
                            newPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.Set(pipeSize.Value); // thiết lập kích thước cho ống mới
                        }
                        newPipe.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM)?.Set(Utils.UnitUtils.MmToFeet(offset));
                        doc.Delete(pipe.Id);
                        newCurves.Add(newPipe);
                    }
                }

                foreach (var duct in mepCurves.OfType<Duct>())
                {
                    if (duct == null) continue;
                    XYZ startPoint = (duct.Location as LocationCurve)?.Curve.GetEndPoint(0);
                    XYZ endPoint = (duct.Location as LocationCurve)?.Curve.GetEndPoint(1);
                    var systemId = duct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();
                    var ductSize = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsDouble(); // lấy kích thước
                    double offset = offsets.ContainsKey(systemId) ? offsets[systemId] : 2800; // Default offset if not provided
                    if (startPoint != null && endPoint != null)
                    {
                        var newDuct = Duct.Create(doc, systemId, duct.GetTypeId(), level.Id, startPoint, endPoint);
                        if (ductSize.HasValue)
                        {
                            newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.Set(ductSize.Value); // thiết lập kích thước cho ống dẫn mới
                        }
                        newDuct.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM)?.Set(Utils.UnitUtils.MmToFeet(offset));
                        doc.Delete(duct.Id);
                        newCurves.Add(newDuct);
                    }
                }

                List<XYZ> xPoints;
                List<dynamic> groupedPoints;

                // Chia các MEP curves tại các điểm giao cắt và tạo fittings
                try
                {
                    xPoints = MEPCurveUtils.FindIntersectionPoints(newCurves).ToList();
                    List<CustomCurve> customCurves = new List<CustomCurve>();
                    foreach (var mepCurve in newCurves)
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
                    groupedPoints = allXPointsConnectors
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
                        .ToList<dynamic>();
                }
                catch
                {
                    //skip
                    groupedPoints = new List<dynamic>();
                }

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

        private Dictionary<ElementId, double> GetSystemOffsets(Document doc, List<Element> selectedElements)
        {
            var offsets = new Dictionary<ElementId, double>();

            var relevantSystems = new HashSet<ElementId>();

            var placeholders = selectedElements
                .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PlaceHolderPipes || e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PlaceHolderDucts)
                .OfType<MEPCurve>()
                .ToList();

            foreach (var mepCurve in placeholders)
            {
                ElementId systemId;
                if (mepCurve is Duct)
                {
                    systemId = mepCurve.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();
                }
                else
                {
                    systemId = mepCurve.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
                }
                relevantSystems.Add(systemId);
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
                foreach (var systemId in relevantSystems)
                {
                    var systemName = GetSystemName(doc, systemId); // Helper method to get system name
                    tableLayoutPanel.Controls.Add(new Label() { Text = systemName, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, rowIndex);

                    // Get the current offset value for the system
                    double currentOffset = GetCurrentOffsetForSystem(doc, systemId, placeholders);

                    TextBox textBox = new TextBox() { Tag = systemId, Text = currentOffset.ToString() }; // Set default value to current offset
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
                            offsets[(ElementId)textBox.Tag] = offset;
                        }
                    }
                }
            }

            return offsets;
        }

        private double GetCurrentOffsetForSystem(Document doc, ElementId systemId, List<MEPCurve> mepCurves)
        {
            // Find an example MEP curve for the system to get its current offset
            foreach (var mepCurve in mepCurves)
            {
                ElementId currentSystemId;
                if (mepCurve is Duct)
                {
                    currentSystemId = mepCurve.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();
                }
                else
                {
                    currentSystemId = mepCurve.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
                }

                if (currentSystemId == systemId)
                {
                    double? offset = mepCurve.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM)?.AsDouble();
                    if (offset.HasValue)
                    {
                        return UnitUtils.FeetToMm(offset.Value); // Convert from feet to mm
                    }
                }
            }
            return 2800; // Default value if no offset found
        }

        private string GetSystemName(Document doc, ElementId systemId)
        {
            var systemElement = doc.GetElement(systemId);
            return systemElement.Name;
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
