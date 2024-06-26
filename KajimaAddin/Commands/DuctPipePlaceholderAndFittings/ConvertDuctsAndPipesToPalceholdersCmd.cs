using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using DocumentFormat.OpenXml.Wordprocessing;
using SKToolsAddins.Utils;
using Document = Autodesk.Revit.DB.Document;
using Form = System.Windows.Forms.Form;
using Level = Autodesk.Revit.DB.Level;
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

            var selectedElements = uidoc.Selection.GetElementIds();
            var mepCurves = new List<MEPCurve>();

            if (selectedElements.Count > 0)
            {
                // Process only selected elements
                var selectedDucts = selectedElements
                    .Select(id => doc.GetElement(id))
                    .OfType<Duct>()
                    .Where(d => d.ReferenceLevel.Id == level.Id)
                    .ToList();
                mepCurves.AddRange(selectedDucts);

                var selectedPipes = selectedElements
                    .Select(id => doc.GetElement(id))
                    .OfType<Pipe>()
                    .Where(d => d.ReferenceLevel.Id == level.Id)
                    .ToList();
                mepCurves.AddRange(selectedPipes);
            }
            else
            {
                // Process all elements in the active view
                var ducts = new FilteredElementCollector(doc, uidoc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_DuctCurves)
                    .OfClass(typeof(Duct))
                    .WhereElementIsNotElementType()
                    .Cast<Duct>()
                    .Where(d => d.ReferenceLevel.Id == level.Id)
                    .ToList();
                mepCurves.AddRange(ducts);

                var pipes = new FilteredElementCollector(doc, uidoc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_PipeCurves)
                    .OfClass(typeof(Pipe))
                    .WhereElementIsNotElementType()
                    .Cast<Pipe>()
                    .Where(d => d.ReferenceLevel.Id == level.Id)
                    .ToList();
                mepCurves.AddRange(pipes);
            }

            // Hiển thị hộp thoại nhập để lấy cao độ system
            var offsets = GetSystemOffsets(doc, mepCurves);

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Convert Pipes and Ducts to Placeholders");
                var ductFittings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DuctFitting)
                    .OfClass(typeof(FamilyInstance))
                    .ToElements()
                    .Cast<FamilyInstance>()
                    .ToList();
                foreach (var ductFitting in ductFittings)
                {
                    var ductFittingName = ductFitting.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString();
                    ductFitting.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ductFittingName);
                }

                // Lưu trữ thông tin hệ thống của fittings
                var pipeFittings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeFitting)
                    .OfClass(typeof(FamilyInstance))
                    .ToElements()
                    .Cast<FamilyInstance>()
                    .ToList();
                foreach (var pipeFitting in pipeFittings)
                {
                    var pipeFittingName = pipeFitting.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                    pipeFitting.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(pipeFittingName);

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
                        var newDuctPlaceholder = Duct.CreatePlaceholder(doc, systemId, duct.GetTypeId(), level.Id, startPoint, endPoint);
                        if (ductSize.HasValue)
                        {
                            newDuctPlaceholder.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.Set(ductSize.Value); // thiết lập kích thước cho placeholder
                        }
                        newDuctPlaceholder.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM)?.Set(UnitUtils.MmToFeet(offset));
                        doc.Delete(duct.Id);
                    }
                }
                
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

        private Dictionary<ElementId, double> GetSystemOffsets(Document doc, List<MEPCurve> mepCurves)
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
                    var systemName = GetSystemName(doc, systemId); // Helper method to get system name
                    tableLayoutPanel.Controls.Add(new Label() { Text = systemName, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, rowIndex);

                    // Get the current offset value for the system
                    double currentOffset = GetCurrentOffsetForSystem(doc, systemId, mepCurves);

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
    }
}
