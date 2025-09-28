using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Control = System.Windows.Forms.Control;
using Form = System.Windows.Forms.Form;

namespace SKRevitAddins.TagTools
{
    // =============== Helpers ===============
    internal static class UnitsUtil
    {
        public static double MmToFt(double mm) =>
            UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);
    }

    /// <summary>
    /// Adapter thống nhất cho TẤT CẢ các loại tag có thể chỉnh:
    /// - IndependentTag: dùng TagHeadPosition get/set (ưu tiên, không cần LocationPoint)
    /// - Bất kỳ phần tử nào có LocationPoint (ví dụ RoomTag/SpatialElementTag, một số tag MEP…)
    /// </summary>
    internal class TagHandle
    {
        public Element Raw { get; }
        public bool Pinned => Raw.Pinned;

        private readonly Func<XYZ> _get;
        private readonly Action<XYZ> _set;

        private TagHandle(Element e, Func<XYZ> getter, Action<XYZ> setter)
        {
            Raw = e;
            _get = getter;
            _set = setter;
        }

        public XYZ GetHead() => _get?.Invoke() ?? XYZ.Zero;

        public void MoveHead(XYZ toPoint)
        {
            if (Pinned) return;
            XYZ cur = GetHead();
            if (toPoint == null) return;
            if (Math.Abs(toPoint.X - cur.X) < 1e-9 &&
                Math.Abs(toPoint.Y - cur.Y) < 1e-9 &&
                Math.Abs(toPoint.Z - cur.Z) < 1e-9) return;

            _set?.Invoke(toPoint);
        }

        // Nhận dạng rộng: cứ là "Tag" (theo class/category) và ta lấy được đầu tag thì dùng được
        public static bool IsTagLike(Element e)
        {
            if (e == null) return false;
            if (e is IndependentTag || e is SpatialElementTag) return true;

            // theo tên lớp
            string tn = e.GetType().Name ?? "";
            if (tn.IndexOf("Tag", StringComparison.OrdinalIgnoreCase) >= 0) return true;

            // theo Category
            var cat = e.Category;
            if (cat != null)
            {
                string cn = cat.Name ?? "";
                if (cn.IndexOf("Tag", StringComparison.OrdinalIgnoreCase) >= 0) return true;

                try
                {
                    var bic = (BuiltInCategory)cat.Id.IntegerValue;
                    if (bic.ToString().IndexOf("Tag", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                }
                catch { /* ignore */ }
            }
            return false;
        }

        public static TagHandle TryCreate(Element e)
        {
            if (!IsTagLike(e)) return null;

            // 1) IndependentTag: dùng TagHeadPosition (không phụ thuộc LocationPoint)
            if (e is IndependentTag it)
            {
                return new TagHandle(
                    e,
                    getter: () => it.TagHeadPosition,
                    setter: p => it.TagHeadPosition = p
                );
            }

            // 2) Bất kỳ phần tử nào có LocationPoint (SpatialElementTag, một số *Tag khác)
            if (e.Location is LocationPoint lp)
            {
                return new TagHandle(
                    e,
                    getter: () => lp.Point,
                    setter: p =>
                    {
                        XYZ v = p - lp.Point;
                        if (Math.Abs(v.X) > 1e-9 || Math.Abs(v.Y) > 1e-9 || Math.Abs(v.Z) > 1e-9)
                            e.Location.Move(v);
                    }
                );
            }

            // 3) Không lấy được đầu tag ⇒ bỏ
            return null;
        }
    }

    internal class TagSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element e) => TagHandle.IsTagLike(e);
        public bool AllowReference(Reference r, XYZ p) => true;
    }

    internal enum AxisMode { Horizontal, Vertical }
    internal enum MainMode { Align, Reflow }

    // =============== UI ===============
    internal class ModeForm : Form
    {
        private RadioButton rbAlign, rbReflow;
        private Button ok, cancel;
        public MainMode Mode { get; private set; } = MainMode.Align;

        public ModeForm()
        {
            Text = "Shinken Group® - Tag Tools";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = MinimizeBox = false;
            Width = 320; Height = 170;

            rbAlign = new RadioButton { Left = 20, Top = 20, Width = 260, Text = "Align: Căn & giãn đều (ngang/dọc)", Checked = true };
            rbReflow = new RadioButton { Left = 20, Top = 45, Width = 260, Text = "Reflow: Đổi bố trí ngang ↔ dọc" };

            ok = new Button { Text = "Tiếp tục", Left = 90, Width = 90, Top = 80, DialogResult = DialogResult.OK };
            cancel = new Button { Text = "Hủy", Left = 190, Width = 90, Top = 80, DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { rbAlign, rbReflow, ok, cancel });
            AcceptButton = ok; CancelButton = cancel;
        }
        public DialogResult ShowDialogGet(out MainMode mode)
        {
            var dr = ShowDialog();
            mode = rbAlign.Checked ? MainMode.Align : MainMode.Reflow;
            return dr;
        }
    }

    internal class AlignForm : Form
    {
        private RadioButton rbH, rbV;
        private NumericUpDown nud;
        private CheckBox chkFirst;
        private Button ok, cancel;

        public AlignForm()
        {
            Text = "Align Tags";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = MinimizeBox = false;
            Width = 320; Height = 210;

            rbH = new RadioButton { Text = "Ngang (cùng Y)", Checked = true, Left = 20, Top = 15, Width = 150 };
            rbV = new RadioButton { Text = "Dọc (cùng X)", Left = 170, Top = 15, Width = 120 };

            var lbl = new Label { Text = "Khoảng cách (mm):", Left = 20, Top = 55, Width = 140 };
            nud = new NumericUpDown { Left = 170, Top = 50, Width = 100, Minimum = 1, Maximum = 100000, DecimalPlaces = 1, Value = 300 };

            // ✅ Mặc định bật
            chkFirst = new CheckBox { Left = 20, Top = 85, Width = 260, Text = "Neo theo tag đầu (thay vì trung bình)", Checked = true };

            ok = new Button { Text = "OK", Left = 90, Width = 90, Top = 120, DialogResult = DialogResult.OK };
            cancel = new Button { Text = "Hủy", Left = 190, Width = 90, Top = 120, DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { rbH, rbV, lbl, nud, chkFirst, ok, cancel });
            AcceptButton = ok; CancelButton = cancel;
        }

        public DialogResult ShowDialogGet(out AxisMode axis, out double spacingMm, out bool anchorFirst)
        {
            var dr = ShowDialog();
            axis = rbH.Checked ? AxisMode.Horizontal : AxisMode.Vertical;
            spacingMm = (double)nud.Value;
            anchorFirst = chkFirst.Checked;   // sẽ là true theo mặc định
            return dr;
        }
    }


    internal class ReflowForm : Form
    {
        private NumericUpDown nud;
        private CheckBox chkMin;
        private Button ok, cancel;

        public ReflowForm()
        {
            Text = "Reflow Tags (Ngang ↔ Dọc)";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = MinimizeBox = false;
            Width = 320; Height = 170;

            var lbl = new Label { Text = "Khoảng cách (mm):", Left = 20, Top = 20, Width = 140 };
            nud = new NumericUpDown { Left = 170, Top = 15, Width = 100, Minimum = 1, Maximum = 100000, DecimalPlaces = 1, Value = 300 };
            chkMin = new CheckBox { Left = 20, Top = 50, Width = 260, Text = "Lấy góc min (trái-dưới) làm mốc" };

            ok = new Button { Text = "OK", Left = 90, Width = 90, Top = 85, DialogResult = DialogResult.OK };
            cancel = new Button { Text = "Hủy", Left = 190, Width = 90, Top = 85, DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { lbl, nud, chkMin, ok, cancel });
            AcceptButton = ok; CancelButton = cancel;
        }

        public DialogResult ShowDialogGet(out double spacingMm, out bool useMinCorner)
        {
            var dr = ShowDialog();
            spacingMm = (double)nud.Value;
            useMinCorner = chkMin.Checked;
            return dr;
        }
    }

    // =============== ONE COMMAND ===============
    [Transaction(TransactionMode.Manual)]
    public class TagToolsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string msg, ElementSet set)
        {
            var ui = c.Application.ActiveUIDocument;
            var doc = ui.Document;

            // 1) Chỉ làm với selection
            var selIds = ui.Selection.GetElementIds();
            if (selIds == null || selIds.Count == 0)
            {
                TaskDialog.Show("Tag Tools", "Hãy chọn các Tag trước khi chạy lệnh.");
                return Result.Cancelled;
            }

            var picked = selIds.Select(id => doc.GetElement(id)).Where(e => e != null).ToList();
            var tags = picked.Select(TagHandle.TryCreate).Where(h => h != null).ToList();

            if (tags.Count == 0)
            {
                TaskDialog.Show("Tag Tools",
                    $"Bạn đã chọn {picked.Count} phần tử nhưng không phần tử nào là Tag có thể chỉnh.\n" +
                    $"- Mẹo: click vào vùng chữ/đầu tag (không phải biểu tượng family).");
                return Result.Cancelled;
            }

            // 2) Chọn chế độ
            var modeForm = new ModeForm();
            if (modeForm.ShowDialogGet(out MainMode mode) != DialogResult.OK) return Result.Cancelled;

            if (mode == MainMode.Align)
            {
                var f = new AlignForm();
                if (f.ShowDialogGet(out AxisMode axis, out double spacingMm, out bool anchorFirst) != DialogResult.OK)
                    return Result.Cancelled;

                double spacing = UnitsUtil.MmToFt(spacingMm);
                var list = tags.Select(t => new { Tag = t, P = t.GetHead() })
                               .OrderBy(x => axis == AxisMode.Horizontal ? x.P.X : x.P.Y)
                               .ToList();
                if (list.Count < 2) return Result.Cancelled;

                XYZ anchor = anchorFirst
                    ? list.First().P
                    : new XYZ(
                        axis == AxisMode.Vertical ? list.Average(x => x.P.X) : list.First().P.X,
                        axis == AxisMode.Horizontal ? list.Average(x => x.P.Y) : list.First().P.Y,
                        list.First().P.Z);

                using (var t = new Transaction(doc, "Tag Tools · Align"))
                {
                    t.Start();
                    for (int i = 0; i < list.Count; i++)
                    {
                        var cur = list[i];
                        XYZ to = (axis == AxisMode.Horizontal)
                            ? new XYZ(anchor.X + i * spacing, anchor.Y, cur.P.Z)
                            : new XYZ(anchor.X, anchor.Y + i * spacing, cur.P.Z);
                        cur.Tag.MoveHead(to);
                    }
                    t.Commit();
                }

                TaskDialog.Show("Tag Tools", $"Align {list.Count} tag theo {(axis == AxisMode.Horizontal ? "NGANG" : "DỌC")} (spacing {spacingMm} mm).");
                return Result.Succeeded;
            }
            else // Reflow
            {
                var f = new ReflowForm();
                if (f.ShowDialogGet(out double spacingMm, out bool useMinCorner) != DialogResult.OK)
                    return Result.Cancelled;

                double spacing = UnitsUtil.MmToFt(spacingMm);
                var pts = tags.Select(t => t.GetHead()).ToList();
                double spreadX = pts.Max(p => p.X) - pts.Min(p => p.X);
                double spreadY = pts.Max(p => p.Y) - pts.Min(p => p.Y);
                bool isHorizontalNow = spreadX >= spreadY;

                var ordered = isHorizontalNow
                    ? tags.OrderBy(t => t.GetHead().Y).ThenBy(t => t.GetHead().X).ToList()
                    : tags.OrderBy(t => t.GetHead().X).ThenBy(t => t.GetHead().Y).ToList();

                XYZ anchor = useMinCorner
                    ? new XYZ(pts.Min(p => p.X), pts.Min(p => p.Y), pts.First().Z)
                    : ordered.First().GetHead();

                using (var t = new Transaction(doc, "Tag Tools · Reflow"))
                {
                    t.Start();
                    for (int i = 0; i < ordered.Count; i++)
                    {
                        var th = ordered[i];
                        XYZ cur = th.GetHead();
                        XYZ to = isHorizontalNow
                            ? new XYZ(anchor.X, anchor.Y + i * spacing, cur.Z)     // sang dọc
                            : new XYZ(anchor.X + i * spacing, anchor.Y, cur.Z);     // sang ngang
                        th.MoveHead(to);
                    }
                    t.Commit();
                }

                TaskDialog.Show("Tag Tools",
                    $"Reflow {tags.Count} tag: {(isHorizontalNow ? "NGANG → DỌC" : "DỌC → NGANG")} (spacing {spacingMm} mm).");
                return Result.Succeeded;
            }
        }
    }
}
