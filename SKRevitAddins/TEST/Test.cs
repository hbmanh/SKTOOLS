using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Form = System.Windows.Forms.Form;
using Point = System.Drawing.Point;

namespace RevitParameterEditor
{
    [Transaction(TransactionMode.Manual)]
    public class ParameterEditorCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // List of target parameters from the image
            List<string> targetParameters = new List<string>
            {
                "Civil Firm-SP",
                "Civil Engineer of Record-SP",
                "Civil Firm Phone Number-SP",
                "Civil Firm Address-SP"
            };

            try
            {
                // Find all elements with any of the target parameters
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                IList<Element> allElements = collector.WhereElementIsNotElementType().ToElements();

                List<Element> elementsWithParameters = new List<Element>();

                foreach (Element elem in allElements)
                {
                    bool hasAnyParameter = false;
                    foreach (string paramName in targetParameters)
                    {
                        if (elem.LookupParameter(paramName) != null)
                        {
                            hasAnyParameter = true;
                            break;
                        }
                    }

                    if (hasAnyParameter)
                    {
                        elementsWithParameters.Add(elem);
                    }
                }

                if (elementsWithParameters.Count == 0)
                {
                    TaskDialog.Show("No Elements Found", "No elements with the specified parameters were found in the model.");
                    return Result.Failed;
                }

                // Show form to edit parameters
                using (ParameterEditorForm form = new ParameterEditorForm(doc, elementsWithParameters, targetParameters))
                {
                    form.ShowDialog();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    public class ParameterEditorForm : Form
    {
        private Document _doc;
        private List<Element> _elements;
        private List<string> _parameterNames;

        private ListBox listElements;
        private DataGridView gridParameters;
        private Button btnApply;
        private Button btnCancel;
        private Label lblElementCount;

        public ParameterEditorForm(Document doc, List<Element> elements, List<string> parameterNames)
        {
            _doc = doc;
            _elements = elements;
            _parameterNames = parameterNames;

            InitializeComponents();
            PopulateElementsList();
        }

        private void InitializeComponents()
        {
            this.Text = "Parameter Editor";
            this.Size = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(900, 600);

            // Element list 
            Label lblElements = new Label
            {
                Text = "Elements with Parameters:",
                Location = new Point(12, 15),
                AutoSize = true
            };
            this.Controls.Add(lblElements);

            listElements = new ListBox
            {
                Location = new Point(12, 40),
                Size = new Size(250, 470),
                SelectionMode = SelectionMode.One
            };
            listElements.SelectedIndexChanged += ListElements_SelectedIndexChanged;
            this.Controls.Add(listElements);

            lblElementCount = new Label
            {
                Text = "Found: 0 elements",
                Location = new Point(12, 520),
                AutoSize = true
            };
            this.Controls.Add(lblElementCount);

            // Parameters grid
            Label lblParameters = new Label
            {
                Text = "Parameters:",
                Location = new Point(280, 15),
                AutoSize = true
            };
            this.Controls.Add(lblParameters);

            gridParameters = new DataGridView
            {
                Location = new Point(280, 40),
                Size = new Size(590, 470),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                EditMode = DataGridViewEditMode.EditOnEnter
            };

            // Create columns for the grid
            DataGridViewTextBoxColumn colParameterName = new DataGridViewTextBoxColumn
            {
                Name = "ParameterName",
                HeaderText = "Parameter Name",
                ReadOnly = true
            };

            DataGridViewCheckBoxColumn colSpaces = new DataGridViewCheckBoxColumn
            {
                Name = "Spaces",
                HeaderText = "Spaces",
                Width = 60
            };

            DataGridViewTextBoxColumn colPrefix = new DataGridViewTextBoxColumn
            {
                Name = "Prefix",
                HeaderText = "Prefix"
            };

            DataGridViewTextBoxColumn colValue = new DataGridViewTextBoxColumn
            {
                Name = "Value",
                HeaderText = "Value"
            };

            DataGridViewTextBoxColumn colSuffix = new DataGridViewTextBoxColumn
            {
                Name = "Suffix",
                HeaderText = "Suffix"
            };

            DataGridViewCheckBoxColumn colBreak = new DataGridViewCheckBoxColumn
            {
                Name = "Break",
                HeaderText = "Break",
                Width = 60
            };

            gridParameters.Columns.AddRange(new DataGridViewColumn[] {
                colParameterName, colSpaces, colPrefix, colValue, colSuffix, colBreak
            });

            this.Controls.Add(gridParameters);

            // Buttons
            btnApply = new Button
            {
                Text = "Apply Changes",
                Location = new Point(750, 520),
                Size = new Size(120, 30),
                Enabled = false
            };
            btnApply.Click += BtnApply_Click;
            this.Controls.Add(btnApply);

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(620, 520),
                Size = new Size(120, 30)
            };
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);
        }

        private void PopulateElementsList()
        {
            listElements.Items.Clear();
            foreach (Element element in _elements)
            {
                string elementName = element.Name;
                if (string.IsNullOrEmpty(elementName))
                {
                    elementName = element.Category?.Name + " " + element.Id.IntegerValue.ToString();
                }

                listElements.Items.Add(new ElementItem(element, elementName));
            }

            lblElementCount.Text = $"Found: {_elements.Count} elements";

            if (listElements.Items.Count > 0)
            {
                listElements.SelectedIndex = 0;
            }
        }

        private void ListElements_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listElements.SelectedItem != null)
            {
                Element selectedElement = ((ElementItem)listElements.SelectedItem).Element;
                PopulateParametersGrid(selectedElement);
                btnApply.Enabled = true;
            }
            else
            {
                gridParameters.Rows.Clear();
                btnApply.Enabled = false;
            }
        }

        private void PopulateParametersGrid(Element element)
        {
            gridParameters.Rows.Clear();

            foreach (string paramName in _parameterNames)
            {
                Parameter param = element.LookupParameter(paramName);
                if (param != null)
                {
                    string value = GetParameterStringValue(param);

                    // Default values based on your image
                    bool hasSpaces = paramName == "Civil Firm-SP" || paramName == "Civil Firm Address-SP";
                    string prefix = "";
                    string suffix = "";
                    bool hasBreak = false;

                    // Set specific values from your image
                    if (paramName == "Civil Firm Phone Number-SP")
                    {
                        prefix = "+81-3-";
                    }
                    else if (paramName == "Civil Firm Address-SP")
                    {
                        prefix = "1-1-1 ";
                    }

                    // Create row with parameter values
                    int rowIndex = gridParameters.Rows.Add();
                    DataGridViewRow row = gridParameters.Rows[rowIndex];

                    // Add parameter details to the row
                    row.Cells["ParameterName"].Value = paramName;
                    row.Cells["Spaces"].Value = hasSpaces;
                    row.Cells["Prefix"].Value = prefix;
                    row.Cells["Value"].Value = value;
                    row.Cells["Suffix"].Value = suffix;
                    row.Cells["Break"].Value = hasBreak;

                    // Store parameter for future use
                    row.Tag = param;
                }
            }
        }

        private string GetParameterStringValue(Parameter param)
        {
            if (param == null) return "";

            switch (param.StorageType)
            {
                case StorageType.String:
                    return param.AsString() ?? "";
                case StorageType.Integer:
                    return param.AsInteger().ToString();
                case StorageType.Double:
                    return param.AsDouble().ToString();
                case StorageType.ElementId:
                    return param.AsElementId().IntegerValue.ToString();
                default:
                    return "";
            }
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            if (listElements.SelectedItem == null) return;

            Element selectedElement = ((ElementItem)listElements.SelectedItem).Element;

            using (Transaction t = new Transaction(_doc, "Update Parameters"))
            {
                t.Start();

                foreach (DataGridViewRow row in gridParameters.Rows)
                {
                    Parameter param = row.Tag as Parameter;
                    if (param != null && !param.IsReadOnly)
                    {
                        // Get values from grid
                        string prefix = row.Cells["Prefix"].Value?.ToString() ?? "";
                        string value = row.Cells["Value"].Value?.ToString() ?? "";
                        string suffix = row.Cells["Suffix"].Value?.ToString() ?? "";

                        // Combine values
                        string combinedValue = prefix + value + suffix;

                        // Set parameter value based on storage type
                        SetParameterValue(param, combinedValue);
                    }
                }

                t.Commit();
            }

            MessageBox.Show("Parameters updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SetParameterValue(Parameter param, string value)
        {
            if (param == null || param.IsReadOnly) return;

            switch (param.StorageType)
            {
                case StorageType.String:
                    param.Set(value);
                    break;
                case StorageType.Integer:
                    if (int.TryParse(value, out int intValue))
                        param.Set(intValue);
                    break;
                case StorageType.Double:
                    if (double.TryParse(value, out double doubleValue))
                        param.Set(doubleValue);
                    break;
                case StorageType.ElementId:
                    if (int.TryParse(value, out int idValue))
                        param.Set(new ElementId(idValue));
                    break;
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }

    // Helper class to store element data in listbox
    public class ElementItem
    {
        public Element Element { get; private set; }
        public string DisplayName { get; private set; }

        public ElementItem(Element element, string displayName)
        {
            Element = element;
            DisplayName = displayName;
        }

        public override string ToString()
        {
            return DisplayName;
        }
    }
}