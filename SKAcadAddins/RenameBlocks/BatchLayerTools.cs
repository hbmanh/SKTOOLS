using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Color = System.Drawing.Color;

namespace BatchLayerTools
{
    public class Commands
    {
        private const string GroupName = "BatchLayerTools";
        // Lệnh chính
        [CommandMethod(GroupName, "BatchRenameLayerInBlocks", CommandFlags.Modal)]
        public void BatchRenameLayerInBlocks()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Chọn BlockReference
            var filter = new SelectionFilter(new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "INSERT")
            });
            var pr = ed.GetSelection(filter);
            if (pr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nKhông có block nào được chọn.");
                return;
            }

            // Thu thập layer của các entity trong block definitions
            var layerNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject sel in pr.Value)
                {
                    if (sel != null)
                    {
                        var bref = tr.GetObject(sel.ObjectId, OpenMode.ForRead) as BlockReference;
                        if (bref != null)
                        {
                            var btr = (BlockTableRecord)tr.GetObject(bref.BlockTableRecord, OpenMode.ForRead);
                            foreach (ObjectId id in btr)
                            {
                                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                                if (ent != null)
                                    layerNames.Add(ent.Layer);
                            }
                        }
                    }
                }
                tr.Commit();
            }

            if (layerNames.Count == 0)
            {
                ed.WriteMessage("\nKhông tìm thấy entity trong các block được chọn.");
                return;
            }

            // Lấy tất cả layer trong bản vẽ
            List<string> allLayers;
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(doc.Database.LayerTableId, OpenMode.ForRead);
                allLayers = lt.Cast<ObjectId>()
                    .Select(id => ((LayerTableRecord)tr.GetObject(id, OpenMode.ForRead)).Name)
                    .OrderBy(n => n)
                    .ToList();
            }

            // Hiển thị form
            var form = new LayerBatchRenameForm(doc, layerNames.ToList(), allLayers);
            Application.ShowModalDialog(form);
        }

        // Alias lệnh tắt RB
        [CommandMethod("reblock")]
        public void Reblock() => BatchRenameLayerInBlocks();
    }

    public class LayerBatchRenameForm : Form
    {
        private readonly Document _doc;
        private readonly List<string> _layers;
        private readonly List<string> _allLayers;
        private readonly Dictionary<string, Color> _newColors;
        private DataGridView dataGridView1;
        private Button btnApply;
        private Button btnCancel;

        public LayerBatchRenameForm(Document doc, List<string> layers, List<string> allLayers)
        {
            _doc = doc;
            _layers = layers;
            _allLayers = allLayers;
            _newColors = new Dictionary<string, Color>();
            InitializeComponent();
            PopulateGrid();
        }

        private void InitializeComponent()
        {
            this.dataGridView1 = new DataGridView();
            this.btnApply = new Button();
            this.btnCancel = new Button();

            var colCurrent = new DataGridViewTextBoxColumn
            {
                HeaderText = "Current Layer",
                Name = "colCurrent",
                ReadOnly = true,
                Width = 180
            };
            var colNewLayer = new DataGridViewComboBoxColumn
            {
                HeaderText = "Select New Layer",
                Name = "colNewLayer",
                Width = 200,
                DataSource = _allLayers,
                FlatStyle = FlatStyle.Flat
            };
            var colNewColor = new DataGridViewButtonColumn
            {
                HeaderText = "New Color",
                Name = "colNewColor",
                Text = "Select…",
                UseColumnTextForButtonValue = true,
                Width = 120
            };

            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new DataGridViewColumn[] { colCurrent, colNewLayer, colNewColor });
            this.dataGridView1.Location = new Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new Size(600, 330);
            this.dataGridView1.EditingControlShowing += DataGridView1_EditingControlShowing;
            this.dataGridView1.CellContentClick += DataGridView1_CellContentClick;
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();

            this.btnApply.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.btnApply.Location = new Point(420, 360);
            this.btnApply.Size = new Size(80, 30);
            this.btnApply.Text = "Apply";
            this.btnApply.Click += BtnApply_Click;

            this.btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.btnCancel.Location = new Point(520, 360);
            this.btnCancel.Size = new Size(80, 30);
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += BtnCancel_Click;

            this.ClientSize = new Size(624, 402);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.btnCancel);
            this.Text = "Batch Rename & Color Layers";
        }

        private void PopulateGrid()
        {
            dataGridView1.Rows.Clear();
            foreach (var layer in _layers.OrderBy(x => x))
            {
                int idx = dataGridView1.Rows.Add();
                var row = dataGridView1.Rows[idx];
                row.Cells["colCurrent"].Value = layer;
                ((DataGridViewComboBoxCell)row.Cells["colNewLayer"]).Value = layer;
                row.Cells["colNewColor"].Value = "Select…";
                _newColors[layer] = Color.Empty;
            }
        }

        private void DataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dataGridView1.CurrentCell.OwningColumn.Name == "colNewLayer" && e.Control is ComboBox combo)
            {
                combo.DropDownStyle = ComboBoxStyle.DropDown;
                combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                combo.AutoCompleteSource = AutoCompleteSource.ListItems;
            }
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "colNewColor" && e.RowIndex >= 0)
            {
                var current = dataGridView1.Rows[e.RowIndex].Cells["colCurrent"].Value.ToString();
                using (var dlg = new ColorDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        _newColors[current] = dlg.Color;
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = dlg.Color;
                    }
                }
            }
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            var ed = _doc.Editor;
            using (var tr = _doc.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(_doc.Database.LayerTableId, OpenMode.ForRead);
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    string oldName = row.Cells["colCurrent"].Value.ToString();
                    string newName = row.Cells["colNewLayer"].Value.ToString();
                    Color newColor = _newColors[oldName];

                    if (!string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase))
                        ed.Command("-LAYER", "R", oldName, newName);

                    if (newColor != Color.Empty && lt.Has(newName))
                    {
                        var ltr = (LayerTableRecord)tr.GetObject(lt[newName], OpenMode.ForWrite);
                        ltr.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(newColor.R, newColor.G, newColor.B);
                    }
                }
                tr.Commit();
            }
            MessageBox.Show("Đã áp dụng thay đổi.", "Batch Layer Rename", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e) => Close();
    }
}
