using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using SKToolsAddins.Utils;
using Form = System.Windows.Forms.Form;
using Panel = System.Windows.Forms.Panel;
using TextBox = System.Windows.Forms.TextBox;
using UnitUtils = SKToolsAddins.Utils.UnitUtils;

namespace SKToolsAddins.Commands.DuctPipePlaceholderAndFittings
{
    [Transaction(TransactionMode.Manual)]
    public class ConvertDuctsAndPipesToPlaceholdersCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            Level level = uidoc.ActiveView.GenLevel;

            var mepCurves = new List<MEPCurve>();

            // Thu thập ducts từ level hiện tại
            var ducts = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctCurves)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .Cast<Duct>()
                .Where(d => d.ReferenceLevel.Id == level.Id)
                .ToList();
            mepCurves.AddRange(ducts);

            // Thu thập pipes từ level hiện tại
            var pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .OfClass(typeof(Pipe))
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .Where(d => d.ReferenceLevel.Id == level.Id)
                .ToList();
            mepCurves.AddRange(pipes);

            // Hiển thị hộp thoại nhập để lấy cao độ cho các hệ thống
            var offsets = GetSystemOffsets(mepCurves);

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Convert Pipes and Ducts to Placeholders");

                foreach (var duct in ducts)
                {
                    if (duct == null) continue;
                    XYZ startPoint = (duct.Location as LocationCurve)?.Curve.GetEndPoint(0);
                    XYZ endPoint = (duct.Location as LocationCurve)?.Curve.GetEndPoint(1);
                    var systemId = duct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();
                    var ductSize = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsDouble(); // lấy kích thước
                    double offset = offsets.ContainsKey(systemId) ? offsets[systemId] : 2800; // Default offset if not provided
                    if (startPoint != null && endPoint != null)
                    {
                        var newDuctPlaceholder = Duct.CreatePlaceholder(doc, systemId, duct.GetTypeId(), level.Id, startPoint, endPoint);
                        if (ductSize.HasValue)
                        {
                            newDuctPlaceholder.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.Set(ductSize.Value); // thiết lập kích thước cho placeholder
                        }
                        newDuctPlaceholder.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM)?.Set(UnitUtils.MmToFeet(offset));
                        doc.Delete(duct.Id);
                    }
                }

                foreach (var pipe in pipes)
                {
                    if (pipe == null) continue;
                    XYZ startPoint = (pipe.Location as LocationCurve)?.Curve.GetEndPoint(0);
                    XYZ endPoint = (pipe.Location as LocationCurve)?.Curve.GetEndPoint(1);
                    var systemId = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
                    var pipeSize = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsDouble(); // lấy kích thước
                    double offset = offsets.ContainsKey(systemId) ? offsets[systemId] : 2800; // Default offset if not provided
                    if (startPoint != null && endPoint != null)
                    {
                        var newPipePlaceholder = Pipe.CreatePlaceholder(doc, systemId, pipe.GetTypeId(), level.Id, startPoint, endPoint);
                        if (pipeSize.HasValue)
                        {
                            newPipePlaceholder.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.Set(pipeSize.Value); // thiết lập kích thước cho placeholder
                        }
                        newPipePlaceholder.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM)?.Set(UnitUtils.MmToFeet(offset));
                        doc.Delete(pipe.Id);
                    }
                }

                trans.Commit();
            }

            TaskDialog.Show("Success", "Successfully converted Pipes and Ducts to Placeholders and updated fittings.");
            return Result.Succeeded;
        }

        private Dictionary<ElementId, double> GetSystemOffsets(List<MEPCurve> mepCurves)
        {
            var offsets = new Dictionary<ElementId, double>();

            var relevantSystems = new HashSet<ElementId>();
            foreach (var mepCurve in mepCurves)
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
                    var systemName = GetSystemName(systemId); // Helper method to get system name
                    tableLayoutPanel.Controls.Add(new Label() { Text = systemName, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, rowIndex);
                    TextBox textBox = new TextBox() { Tag = systemId, Text = "2800" }; // Set default value to 2800
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

        private string GetSystemName(ElementId systemId)
        {
            // Implement a method to get the system name based on the systemId
            // This can be a lookup in the document's element list or a predefined dictionary
            // For example:
            // var systemElement = doc.GetElement(systemId);
            // return systemElement.Name;
            return systemId.ToString(); // Replace with actual implementation
        }
    }
}
