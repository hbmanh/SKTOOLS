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

// Đăng ký lệnh để AutoCAD nhận diện
[assembly: CommandClass(typeof(BatchLayerTools.Commands))]

namespace BatchLayerTools
{
    public class Commands
    {
        private const string GroupName = "BatchLayerTools";

        // Lệnh chính
        [CommandMethod(GroupName, "BatchRenameLayerInBlocks", CommandFlags.Modal)]
        public void BatchRenameLayerInBlocks()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

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
            var blockRefs = new List<ObjectId>();
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject sel in pr.Value)
                {
                    if (sel == null) continue;
                    var bref = tr.GetObject(sel.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (bref == null) continue;
                    blockRefs.Add(sel.ObjectId);
                    var btr = tr.GetObject(bref.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    foreach (ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent != null)
                            layerNames.Add(ent.Layer);
                    }
                    layerNames.Add(bref.Layer); // layer của chính block reference
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
                tr.Commit();
            }

            // Hiển thị form
            var form = new LayerBatchRenameForm(doc, blockRefs, layerNames.ToList(), allLayers);
            Application.ShowModalDialog(form);
        }

        // Alias lệnh tắt REBLOCK
        [CommandMethod(GroupName, "REBLOCK", CommandFlags.Modal)]
        public void Reblock() => BatchRenameLayerInBlocks();
    }

    public class LayerBatchRenameForm : Form
    {
        private readonly Document _doc;
        private readonly List<ObjectId> _selectedBlockRefs;
        private readonly List<string> _layers;
        private readonly List<string> _allLayers;
        private CheckBox chkChangeBlockItself;
        private CheckBox chkChangeInsideEntities;
        private DataGridView dataGridView1;
        private ComboBox cboBlockTargetLayer;
        private Label lblBlockLayer;
        private Button btnApply;
        private Button btnCancel;

        public LayerBatchRenameForm(Document doc, List<ObjectId> selectedBlockRefs, List<string> layers, List<string> allLayers)
        {
            _doc = doc;
            _selectedBlockRefs = selectedBlockRefs;
            _layers = layers;
            _allLayers = allLayers;
            InitializeComponent();
            PopulateGrid();
        }

        private void InitializeComponent()
        {
            this.dataGridView1 = new DataGridView();
            this.btnApply = new Button();
            this.btnCancel = new Button();
            this.chkChangeBlockItself = new CheckBox();
            this.chkChangeInsideEntities = new CheckBox();
            this.cboBlockTargetLayer = new ComboBox();
            this.lblBlockLayer = new Label();

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
            var colNewColor = new DataGridViewComboBoxColumn
            {
                HeaderText = "New Color (ACI)",
                Name = "colNewColor",
                Width = 120,
                FlatStyle = FlatStyle.Flat
            };
            for (int i = 1; i <= 255; i++) colNewColor.Items.Add(i.ToString());

            this.dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(colCurrent, colNewLayer, colNewColor);
            this.dataGridView1.Location = new Point(12, 90);
            this.dataGridView1.Size = new Size(600, 250);

            this.chkChangeBlockItself.Text = "Đổi layer của Block (bên ngoài)";
            this.chkChangeBlockItself.Location = new Point(12, 10);
            this.chkChangeBlockItself.Size = new Size(250, 20);
            this.chkChangeBlockItself.Checked = true;
            this.chkChangeBlockItself.CheckedChanged += (s, e) =>
            {
                cboBlockTargetLayer.Enabled = chkChangeBlockItself.Checked;
                lblBlockLayer.Enabled = chkChangeBlockItself.Checked;
            };

            this.chkChangeInsideEntities.Text = "Đổi layer của đối tượng bên trong Block";
            this.chkChangeInsideEntities.Location = new Point(12, 30);
            this.chkChangeInsideEntities.Size = new Size(300, 20);
            this.chkChangeInsideEntities.Checked = true;

            this.lblBlockLayer.Text = "Layer cho Block:";
            this.lblBlockLayer.Location = new Point(30, 55);
            this.lblBlockLayer.Size = new Size(100, 20);

            this.cboBlockTargetLayer.Location = new Point(140, 52);
            this.cboBlockTargetLayer.Size = new Size(200, 24);
            this.cboBlockTargetLayer.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cboBlockTargetLayer.Items.AddRange(_allLayers.ToArray());
            if (_allLayers.Count > 0)
                this.cboBlockTargetLayer.SelectedIndex = 0;

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
            this.Controls.Add(this.chkChangeBlockItself);
            this.Controls.Add(this.chkChangeInsideEntities);
            this.Controls.Add(this.lblBlockLayer);
            this.Controls.Add(this.cboBlockTargetLayer);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.btnCancel);
            this.Text = "Batch Change Layer in Blocks";
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
                ((DataGridViewComboBoxCell)row.Cells["colNewColor"]).Value = "";
            }
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            using (var tr = _doc.TransactionManager.StartTransaction())
            {
                foreach (ObjectId blockId in _selectedBlockRefs)
                {
                    BlockReference bref = tr.GetObject(blockId, OpenMode.ForWrite) as BlockReference;
                    if (bref == null) continue;

                    // Đổi layer của chính block reference
                    if (chkChangeBlockItself.Checked)
                    {
                        if (cboBlockTargetLayer.SelectedItem != null)
                            bref.Layer = cboBlockTargetLayer.SelectedItem.ToString();
                    }

                    // Đổi layer của entity bên trong block
                    if (chkChangeInsideEntities.Checked)
                    {
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bref.BlockTableRecord, OpenMode.ForRead);
                        foreach (ObjectId entId in btr)
                        {
                            Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                            if (ent == null) continue;

                            foreach (DataGridViewRow row in dataGridView1.Rows)
                            {
                                if (row.Cells["colCurrent"].Value == null || row.Cells["colNewLayer"].Value == null)
                                    continue;
                                string oldLayer = row.Cells["colCurrent"].Value.ToString();
                                string newLayer = row.Cells["colNewLayer"].Value.ToString();
                                string aciStr = row.Cells["colNewColor"].Value?.ToString();
                                short aci;
                                short.TryParse(aciStr, out aci);

                                if (ent.Layer == oldLayer)
                                {
                                    ent.Layer = newLayer;
                                    if (aci >= 1 && aci <= 255)
                                        ent.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, aci);
                                }
                            }
                        }
                    }
                }
                tr.Commit();
            }

            MessageBox.Show("Đã cập nhật layer theo lựa chọn.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e) => Close();
    }
}
