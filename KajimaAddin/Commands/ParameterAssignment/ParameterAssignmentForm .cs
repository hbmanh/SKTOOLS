using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Form = System.Windows.Forms.Form;
using ParamObj = SKToolsAddins.Commands.ParameterAssignment.CreateSharedParamFromExcelCmd.ParamObj;

namespace SKToolsAddins.Commands.ParameterAssignment
{
    public partial class ParameterAssignmentForm : Form
    {
        private Document _doc;
        public List<ParamObj> SelectedParameters { get; private set; }

        public ParameterAssignmentForm(List<ParamObj> parameters, Document doc)
        {
            InitializeComponent();
            _doc = doc;

            // Initialize DataGridView with columns
            dataGridView1.Columns.Add("Parameter", "パラメータ名");
            foreach (var paramObj in parameters)
            {
                foreach (var category in paramObj.Categories)
                {
                    if (!dataGridView1.Columns.Contains(category.Name))
                    {
                        var checkBoxColumn = new DataGridViewCheckBoxColumn
                        {
                            HeaderText = category.Name,
                            Name = category.Name
                        };
                        dataGridView1.Columns.Add(checkBoxColumn);
                    }
                }
            }

            // Add parameters to the DataGridView
            foreach (var paramObj in parameters)
            {
                var row = new DataGridViewRow();
                row.CreateCells(dataGridView1, paramObj.ParamName);

                // Set checkbox values based on the paramCategoryDict
                foreach (var category in paramObj.Categories)
                {
                    int columnIndex = dataGridView1.Columns[category.Name].Index;
                    row.Cells[columnIndex].Value = true;
                }

                dataGridView1.Rows.Add(row);
            }

            // Auto resize DataGridView columns and rows to fit content
            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoResizeRows();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            SelectedParameters = new List<ParamObj>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[0].Value != null)
                {
                    var parameterName = row.Cells[0].Value.ToString();
                    var paramObj = new ParamObj(parameterName, true);

                    for (int i = 1; i < row.Cells.Count; i++)
                    {
                        var cell = row.Cells[i] as DataGridViewCheckBoxCell;
                        if (cell != null && Convert.ToBoolean(cell.Value) == true)
                        {
                            var category = GetCategoryByName(_doc, dataGridView1.Columns[i].HeaderText);
                            if (category != null)
                            {
                                paramObj.Categories.Add(category);
                            }
                        }
                    }

                    SelectedParameters.Add(paramObj);
                }
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private Category GetCategoryByName(Document doc, string categoryName)
        {
            foreach (Category category in doc.Settings.Categories)
            {
                if (category.Name == categoryName)
                {
                    return category;
                }
            }
            return null;
        }
    }
}
