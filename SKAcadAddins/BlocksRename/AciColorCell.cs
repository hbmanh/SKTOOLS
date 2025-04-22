using static System.Windows.Forms.VisualStyles.VisualStyleElement.Header;
using System.Windows.Forms;
using System;
using System.Drawing;
using Autodesk.AutoCAD.Colors;

public class AciColorCell : DataGridViewComboBoxCell
{
    public override Type EditType => typeof(AciColorComboBox);

    public override object Clone()
    {
        var clone = (AciColorCell)base.Clone();
        return clone;
    }
}

public class AciColorComboBox : ComboBox, IDataGridViewEditingControl
{
    public DataGridView EditingControlDataGridView { get; set; }
    public object EditingControlFormattedValue
    {
        get => SelectedItem?.ToString();
        set => SelectedItem = value;
    }

    public int EditingControlRowIndex { get; set; }
    public bool EditingControlValueChanged { get; set; }
    public Cursor EditingPanelCursor => Cursors.Default;
    public bool RepositionEditingControlOnValueChange => false;

    public AciColorComboBox()
    {
        DropDownStyle = ComboBoxStyle.DropDownList;
        DrawMode = DrawMode.OwnerDrawFixed;
        ItemHeight = 18;

        for (short i = 1; i <= 255; i++)
            Items.Add(i);

        DrawItem += (s, e) =>
        {
            if (e.Index < 0 || e.Index >= Items.Count) return;

            short aci = (short)Items[e.Index];
            var acadColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, aci);
            var sysColor = System.Drawing.Color.FromArgb(acadColor.Red, acadColor.Green, acadColor.Blue);

            e.DrawBackground();
            using (Brush brush = new SolidBrush(sysColor))
            {
                e.Graphics.FillRectangle(brush, e.Bounds.X + 2, e.Bounds.Y + 2, 16, 14);
                e.Graphics.DrawRectangle(Pens.Black, e.Bounds.X + 2, e.Bounds.Y + 2, 16, 14);
            }
            e.Graphics.DrawString(aci.ToString(), e.Font, Brushes.Black, e.Bounds.X + 22, e.Bounds.Y + 2);
        };
    }

    public void ApplyCellStyleToEditingControl(DataGridViewCellStyle dataGridViewCellStyle) { }
    public bool EditingControlWantsInputKey(Keys keyData, bool dataGridViewWantsInputKey) => true;
    public object GetEditingControlFormattedValue(DataGridViewDataErrorContexts context) => EditingControlFormattedValue;
    public void PrepareEditingControlForEdit(bool selectAll) { }
}
