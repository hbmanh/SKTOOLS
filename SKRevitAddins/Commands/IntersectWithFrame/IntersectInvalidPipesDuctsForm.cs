using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Form = System.Windows.Forms.Form;

namespace SKRevitAddins.Commands.IntersectWithFrame
{
    public partial class IntersectInvalidPipesDuctsForm : Form
    {
        private UIDocument _uidoc;
        private Document _doc;

        public IntersectInvalidPipesDuctsForm(UIDocument uidoc, List<Tuple<ElementId, string>> invalidElements)
        {
            InitializeComponent();
            _uidoc = uidoc;
            _doc = uidoc.Document;

            // Populate the DataGridView with invalid elements
            foreach (var invalidElement in invalidElements)
            {
                var row = new DataGridViewRow();
                row.CreateCells(dataGridView1);
                row.Cells[0].Value = invalidElement.Item1.ToString();
                row.Cells[1].Value = invalidElement.Item2;
                dataGridView1.Rows.Add(row);
            }
        }

        private void btnReview_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedId = new ElementId(Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[0].Value));
                var element = _doc.GetElement(selectedId);

                if (element != null)
                {
                    // Open 3D view with section box
                    using (Transaction t = new Transaction(_doc, "Create Section Box"))
                    {
                        t.Start();

                        View3D view3D = new FilteredElementCollector(_doc)
                            .OfClass(typeof(View3D))
                            .Cast<View3D>()
                            .FirstOrDefault(v => !v.IsTemplate);

                        if (view3D != null)
                        {
                            BoundingBoxXYZ bbox = element.get_BoundingBox(view3D);
                            if (bbox != null)
                            {
                                view3D.IsSectionBoxActive = true;
                                view3D.SetSectionBox(bbox);
                            }

                            _uidoc.ActiveView = view3D;
                            _uidoc.ShowElements(element);
                        }

                        t.Commit();
                    }
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
