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

[assembly: CommandClass(typeof(BlockLayerManager.Commands))]

namespace BlockLayerManager
{
    public class Commands
    {
        [CommandMethod("REBLOCK", CommandFlags.Modal)]
        public void Reblock()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            var filter = new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "INSERT") });
            var pr = ed.GetSelection(filter);
            if (pr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nKhông có block nào được chọn.");
                return;
            }

            var blockRefs = new List<ObjectId>();
            var layerNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (var tr = doc.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject sel in pr.Value)
                {
                    if (sel?.ObjectId == null) continue;
                    var bref = tr.GetObject(sel.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (bref == null) continue;

                    blockRefs.Add(bref.ObjectId);
                    layerNames.Add(bref.Layer);

                    var btr = (BlockTableRecord)tr.GetObject(bref.BlockTableRecord, OpenMode.ForRead);
                    foreach (ObjectId id in btr)
                    {
                        if (tr.GetObject(id, OpenMode.ForRead) is Entity ent)
                            layerNames.Add(ent.Layer);
                    }
                }
                tr.Commit();
            }

            List<string> allLayers;
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(doc.Database.LayerTableId, OpenMode.ForRead);
                allLayers = lt.Cast<ObjectId>()
                    .Select(id => ((LayerTableRecord)tr.GetObject(id, OpenMode.ForRead)).Name)
                    .OrderBy(n => n).ToList();
                tr.Commit();
            }

            var form = new LayerManagerForm(doc, blockRefs, layerNames.ToList(), allLayers);
            Application.ShowModalDialog(form);
        }
    }

    public class AciColorCell : DataGridViewComboBoxCell
    {
        public override Type EditType => typeof(AciColorComboBox);
    }

    public class AciColorComboBox : ComboBox, IDataGridViewEditingControl
    {
        public AciColorComboBox()
        {
            DropDownStyle = ComboBoxStyle.DropDownList;
            DrawMode = DrawMode.OwnerDrawFixed;
            for (short i = 1; i <= 255; i++) Items.Add(i);
            DrawItem += DrawPreview;
        }

        private void DrawPreview(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= Items.Count) return;

            short aci = Convert.ToInt16(Items[e.Index]);
            var acadColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, aci);
            var sysColor = ColorTranslator.FromOle(acadColor.ColorValue.ToArgb());

            e.DrawBackground();
            using (Brush brush = new SolidBrush(sysColor))
            {
                e.Graphics.FillRectangle(brush, e.Bounds.X + 2, e.Bounds.Y + 2, 16, 14);
                e.Graphics.DrawRectangle(Pens.Black, e.Bounds.X + 2, e.Bounds.Y + 2, 16, 14);
            }

            e.Graphics.DrawString(aci.ToString(), e.Font, Brushes.Black, e.Bounds.X + 22, e.Bounds.Y + 2);
        }

        public DataGridView EditingControlDataGridView { get; set; }
        public object EditingControlFormattedValue { get => SelectedItem?.ToString(); set => SelectedItem = value; }
        public int EditingControlRowIndex { get; set; }
        public bool EditingControlValueChanged { get; set; }
        public Cursor EditingPanelCursor => Cursors.Default;
        public bool RepositionEditingControlOnValueChange => false;

        public void ApplyCellStyleToEditingControl(DataGridViewCellStyle dataGridViewCellStyle) { }
        public bool EditingControlWantsInputKey(Keys key, bool dataGridViewWantsInputKey) => true;
        public object GetEditingControlFormattedValue(DataGridViewDataErrorContexts context) => EditingControlFormattedValue;
        public void PrepareEditingControlForEdit(bool selectAll) { }
    }

    public class LayerManagerForm : Form
    {
        private readonly Document _doc;
        private readonly List<ObjectId> _blockRefs;
        private readonly List<string> _layers;
        private readonly List<string> _allLayers;
        private readonly DataGridView dataGridView1;
        private readonly ComboBox cboBlockTargetLayer;
        private readonly CheckBox chkBlock, chkInside;
        private readonly Button btnApply, btnCancel;
        private readonly Label lblBlockLayer;

        private const string LogoPath = @"C:\ProgramData\Autodesk\Revit\Addins\2023\SKTools.bundle\Contents\Resources\Images\shinken.png";

        public LayerManagerForm(Document doc, List<ObjectId> blockRefs, List<string> layers, List<string> allLayers)
        {
            _doc = doc;
            _blockRefs = blockRefs;
            _layers = layers;
            _allLayers = allLayers;

            Text = "Block's Layer Manager";
            Size = new Size(720, 520);
            StartPosition = FormStartPosition.CenterScreen;

            var headerPanel = new Panel
            {
                Size = new Size(700, 60),
                Location = new Point(10, 10),
                BackColor = System.Drawing.Color.WhiteSmoke
            };

            var logo = new PictureBox
            {
                Size = new Size(40, 40),
                Location = new Point(10, 10),
                SizeMode = PictureBoxSizeMode.StretchImage
            };
            if (System.IO.File.Exists(LogoPath))
                logo.Image = System.Drawing.Image.FromFile(LogoPath);

            var lblTitle = new Label
            {
                Text = "  Block's Layer Manager - Shinken Group®",
                Location = new Point(60, 18),
                AutoSize = true,
                Font = new System.Drawing.Font("Segoe UI", 14, FontStyle.Bold)
            };
            headerPanel.Controls.Add(logo);
            headerPanel.Controls.Add(lblTitle);
            Controls.Add(headerPanel);

            chkBlock = new CheckBox { Text = "Đổi layer của Block (bên ngoài)", Location = new Point(20, 80), Checked = true, AutoSize = true };
            chkInside = new CheckBox { Text = "Đổi layer của đối tượng bên trong Block", Location = new Point(20, 105), Checked = true, AutoSize = true };

            lblBlockLayer = new Label { Text = "Layer cho Block:", Location = new Point(40, 135), AutoSize = true };
            cboBlockTargetLayer = new ComboBox { Location = new Point(160, 132), Size = new Size(200, 24), DropDownStyle = ComboBoxStyle.DropDownList };
            cboBlockTargetLayer.Items.AddRange(_allLayers.ToArray());
            cboBlockTargetLayer.SelectedIndex = 0;

            chkBlock.CheckedChanged += (s, e) =>
            {
                cboBlockTargetLayer.Enabled = chkBlock.Checked;
                lblBlockLayer.Enabled = chkBlock.Checked;
            };

            dataGridView1 = new DataGridView
            {
                Location = new Point(20, 170),
                Size = new Size(660, 250),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                AllowUserToAddRows = false,
                RowHeadersVisible = false
            };

            var colCurrent = new DataGridViewTextBoxColumn { HeaderText = "Current Layer", Name = "colCurrent", ReadOnly = true, Width = 200 };
            var colNewLayer = new DataGridViewComboBoxColumn { HeaderText = "New Layer", Name = "colNewLayer", Width = 220, DataSource = _allLayers, FlatStyle = FlatStyle.Flat };
            var colNewColor = new DataGridViewColumn(new AciColorCell())
            {
                HeaderText = "Color (ACI)",
                Name = "colNewColor",
                Width = 120
            };

            dataGridView1.Columns.AddRange(colCurrent, colNewLayer, colNewColor);

            btnApply = new Button { Text = "Apply", Location = new Point(500, 440), Size = new Size(80, 30), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnApply.Click += BtnApply_Click;

            btnCancel = new Button { Text = "Cancel", Location = new Point(600, 440), Size = new Size(80, 30), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnCancel.Click += (s, e) => Close();

            Controls.AddRange(new Control[] {
                chkBlock, chkInside, lblBlockLayer, cboBlockTargetLayer,
                dataGridView1, btnApply, btnCancel
            });

            PopulateGrid();
        }

        private void PopulateGrid()
        {
            dataGridView1.Rows.Clear();
            foreach (var layer in _layers.OrderBy(x => x))
            {
                int i = dataGridView1.Rows.Add();
                var row = dataGridView1.Rows[i];
                row.Cells["colCurrent"].Value = layer;

                var cb = new DataGridViewComboBoxCell();
                cb.Items.AddRange(_allLayers.ToArray());
                cb.Value = layer;
                row.Cells["colNewLayer"] = cb;

                var aciCell = new AciColorCell();
                for (short j = 1; j <= 255; j++)
                    aciCell.Items.Add(j.ToString());

                aciCell.Value = "";
                row.Cells["colNewColor"] = aciCell;
            }
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            using (var tr = _doc.TransactionManager.StartTransaction())
            {
                foreach (var blockId in _blockRefs)
                {
                    var bref = tr.GetObject(blockId, OpenMode.ForWrite) as BlockReference;
                    if (bref == null) continue;

                    if (chkBlock.Checked && cboBlockTargetLayer.SelectedItem != null)
                        bref.Layer = cboBlockTargetLayer.SelectedItem.ToString();

                    if (chkInside.Checked)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(bref.BlockTableRecord, OpenMode.ForRead);
                        ApplyToEntitiesRecursive(tr, btr);
                    }
                }

                tr.Commit();
            }

            MessageBox.Show("Đã cập nhật layer.", "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }

        private void ApplyToEntitiesRecursive(Transaction tr, BlockTableRecord btr)
        {
            foreach (ObjectId entId in btr)
            {
                var ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                if (ent == null) continue;

                if (ent is BlockReference nestedRef)
                {
                    var nestedBtr = (BlockTableRecord)tr.GetObject(nestedRef.BlockTableRecord, OpenMode.ForRead);
                    ApplyToEntitiesRecursive(tr, nestedBtr);
                }
                else
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Cells["colCurrent"].Value == null || row.Cells["colNewLayer"].Value == null)
                            continue;

                        string oldLayer = row.Cells["colCurrent"].Value.ToString();
                        string newLayer = row.Cells["colNewLayer"].Value.ToString();
                        var colorValue = row.Cells["colNewColor"].Value;

                        if (ent.Layer == oldLayer)
                        {
                            ent.Layer = newLayer;
                            if (colorValue != null && short.TryParse(colorValue.ToString(), out short aci) && aci >= 1 && aci <= 255)
                                ent.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, aci);
                        }
                    }
                }
            }
        }
    }
}
