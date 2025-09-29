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

namespace SKRevitAddins.TAGTools
{
    // ===== Helpers =====
    internal static class UnitsUtil
    {
        public static double MmToFt(double mm) =>
            UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);
    }

    internal enum AxisMode { Horizontal, Vertical }

    /// <summary>
    /// Adapter thống nhất cho MỌI loại tag có thể chỉnh:
    /// - IndependentTag: get/set TagHeadPosition
    /// - Các tag khác có LocationPoint: move bằng Location.Move
    /// </summary>
    internal class TagHandle
    {
        public Element Raw { get; }
        public bool Pinned => Raw.Pinned;

        private readonly Func<XYZ> _getHead;
        private readonly Action<XYZ> _setHead;

        private TagHandle(Element e, Func<XYZ> getter, Action<XYZ> setter)
        { Raw = e; _getHead = getter; _setHead = setter; }

        public XYZ GetHead() => _getHead?.Invoke() ?? XYZ.Zero;

        public void MoveHead(XYZ to)
        {
            if (Pinned) return;
            XYZ cur = GetHead();
            if (Math.Abs(to.X - cur.X) < 1e-9 &&
                Math.Abs(to.Y - cur.Y) < 1e-9 &&
                Math.Abs(to.Z - cur.Z) < 1e-9) return;

            _setHead?.Invoke(to);
        }

        // Nhận dạng rộng: cứ “Tag” (class/category) là chấp nhận
        public static bool IsTagLike(Element e)
        {
            if (e == null) return false;
            if (e is IndependentTag || e is SpatialElementTag) return true;

            string tn = e.GetType().Name ?? "";
            if (tn.IndexOf("Tag", StringComparison.OrdinalIgnoreCase) >= 0) return true;

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
                catch { }
            }
            return false;
        }

        public static TagHandle TryCreate(Element e)
        {
            if (!IsTagLike(e)) return null;

            // 1) IndependentTag: an toàn & chuẩn nhất
            if (e is IndependentTag it)
            {
                return new TagHandle(
                    e,
                    getter: () => it.TagHeadPosition,
                    setter: p => it.TagHeadPosition = p
                );
            }

            // 2) Các tag khác có LocationPoint
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

            return null; // không lấy được đầu tag
        }
    }

    // ===== Modeless UI (Align-only) =====
    internal class MainForm : Form
    {
        private RadioButton rbH, rbV;
        private NumericUpDown nudSpacing;
        private CheckBox chkAnchorFirst;
        private Button btnApply, btnClose;

        public AxisMode AlignAxis => rbH.Checked ? AxisMode.Horizontal : AxisMode.Vertical;
        public double AlignSpacingMm => (double)nudSpacing.Value;
        public bool AlignAnchorFirst => chkAnchorFirst.Checked;

        public event EventHandler ApplyClicked;

        public MainForm()
        {
            Text = "Shinken Group® - Tag Tools";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            Width = 320; Height = 200;

            rbH = new RadioButton { Left = 20, Top = 15, Width = 130, Text = "Ngang (cùng Y)", Checked = true };
            rbV = new RadioButton { Left = 160, Top = 15, Width = 130, Text = "Dọc (cùng X)" };

            var lbl = new Label { Left = 20, Top = 50, Width = 120, Text = "Spacing (mm):" };
            nudSpacing = new NumericUpDown { Left = 160, Top = 45, Width = 100, Minimum = 1, Maximum = 100000, DecimalPlaces = 1, Value = 300 };

            chkAnchorFirst = new CheckBox
            {
                Left = 20,
                Top = 80,
                Width = 260,
                Text = "Neo theo tag đầu (thay vì trung bình)",
                Checked = true // default ✅
            };

            btnApply = new Button { Left = 120, Top = 115, Width = 80, Text = "Apply" };
            btnClose = new Button { Left = 210, Top = 115, Width = 80, Text = "Close" };

            btnApply.Click += (s, e) => ApplyClicked?.Invoke(this, EventArgs.Empty);
            btnClose.Click += (s, e) => Close();

            Controls.AddRange(new Control[] { rbH, rbV, lbl, nudSpacing, chkAnchorFirst, btnApply, btnClose });
        }
    }

    // ===== ExternalEvent handler (Align) =====
    internal class TagToolsExternalHandler : IExternalEventHandler
    {
        private UIDocument _uiDoc;

        public AxisMode AlignAxis { get; set; }
        public double AlignSpacingFt { get; set; }
        public bool AlignAnchorFirst { get; set; }

        public void Init(UIDocument uiDoc) => _uiDoc = uiDoc;

        public void Execute(UIApplication app)
        {
            if (_uiDoc == null) return;
            var doc = _uiDoc.Document;

            // Mỗi lần Apply: dùng selection hiện tại
            var selIds = _uiDoc.Selection.GetElementIds();
            if (selIds == null || selIds.Count == 0)
            {
                TaskDialog.Show("Tag Tools", "Hãy chọn các Tag rồi bấm Apply.");
                return;
            }

            var tags = selIds
                .Select(id => doc.GetElement(id))
                .Where(e => e != null)
                .Select(TagHandle.TryCreate)
                .Where(h => h != null)
                .ToList();

            if (tags.Count < 2)
            {
                TaskDialog.Show("Tag Tools", "Cần ít nhất 2 Tag hợp lệ để Align.");
                return;
            }

            // Sắp xếp và CHỌN ANCHOR sao cho "dọc đi xuống" từ tag cao nhất
            var list = (AlignAxis == AxisMode.Horizontal)
                ? tags.Select(t => new { Tag = t, P = t.GetHead() })
                      .OrderBy(x => x.P.X)                  // trái → phải
                      .ToList()
                : tags.Select(t => new { Tag = t, P = t.GetHead() })
                      .OrderByDescending(x => x.P.Y)        // trên → dưới (top-most first)
                      .ToList();

            XYZ anchor;
            if (AlignAnchorFirst)
            {
                // neo vào phần tử đầu tiên của danh sách đã sắp xếp
                anchor = list.First().P;
            }
            else
            {
                // neo theo trung bình theo trục vuông góc
                anchor = new XYZ(
                    AlignAxis == AxisMode.Vertical ? list.Average(x => x.P.X) : list.First().P.X,
                    AlignAxis == AxisMode.Horizontal ? list.Average(x => x.P.Y) : list.First().P.Y,
                    list.First().P.Z);
            }

            using (var t = new Transaction(doc, "Tag Tools · Align"))
            {
                t.Start();
                for (int i = 0; i < list.Count; i++)
                {
                    var cur = list[i];
                    XYZ to = (AlignAxis == AxisMode.Horizontal)
                        ? new XYZ(anchor.X + i * AlignSpacingFt, anchor.Y, cur.P.Z)          // ngang: +X, giữ Y
                        : new XYZ(anchor.X, anchor.Y - i * AlignSpacingFt, cur.P.Z);         // dọc: từ TOP đi xuống (−Y), giữ X
                    cur.Tag.MoveHead(to);
                }
                t.Commit();
            }
        }

        public string GetName() => "SK Tag Tools Align";
    }

    // ===== Entry command: show modeless form & wire ExternalEvent =====
    [Transaction(TransactionMode.Manual)]
    public class TagToolsCmd : IExternalCommand
    {
        private static MainForm _form;
        private static ExternalEvent _extEvent;
        private static TagToolsExternalHandler _handler;

        public Result Execute(ExternalCommandData c, ref string msg, ElementSet set)
        {
            var ui = c.Application.ActiveUIDocument;

            if (_form == null || _form.IsDisposed)
            {
                _handler = new TagToolsExternalHandler();
                _handler.Init(ui);
                _extEvent = ExternalEvent.Create(_handler);

                _form = new MainForm();
                _form.ApplyClicked += (s, e) =>
                {
                    _handler.AlignAxis = _form.AlignAxis;
                    _handler.AlignSpacingFt = UnitsUtil.MmToFt(_form.AlignSpacingMm);
                    _handler.AlignAnchorFirst = _form.AlignAnchorFirst;

                    _handler.Init(ui); // refresh phòng khi active doc thay đổi
                    _extEvent.Raise();
                };

                // Show modeless; set owner bằng handle cửa sổ Revit
                var hwnd = c.Application.MainWindowHandle;
                if (hwnd != IntPtr.Zero)
                    _form.Show(new RevitWindowHandle(hwnd));
                else
                    _form.Show();
            }
            else
            {
                _form.Activate();
            }

            return Result.Succeeded;
        }
    }

    // Owner window wrapper
    internal class RevitWindowHandle : IWin32Window
    {
        public IntPtr Handle { get; private set; }
        public RevitWindowHandle(IntPtr h) { Handle = h; }
    }
}
