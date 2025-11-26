using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Reflection;
using View = Autodesk.Revit.DB.View;
using Form = System.Windows.Forms.Form;
using Panel = System.Windows.Forms.Panel;
using Point = System.Drawing.Point;
using Color = System.Drawing.Color;
using CheckBox = System.Windows.Forms.CheckBox;

namespace SKRevitAddins.GENTools
{
    [Transaction(TransactionMode.Manual)]
    public class DuplicateElementsDetectorCmd : IExternalCommand
    {
        public static DuplicateResultForm _openedForm;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            if (_openedForm != null && !_openedForm.IsDisposed)
            {
                _openedForm.BringToFront();
                return Result.Succeeded;
            }

            // 1. INPUT: Ưu tiên Selection, nếu không thì lấy Active View
            IList<Element> elementsToProcess = new List<Element>();
            ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();

            if (selectedIds.Count > 0)
            {
                foreach (ElementId id in selectedIds)
                {
                    Element e = doc.GetElement(id);
                    // Lấy FamilyInstance và cả các đối tượng có Geometry (nếu cần mở rộng sau này)
                    if (e is FamilyInstance && e.Category != null)
                    {
                        elementsToProcess.Add(e);
                    }
                }

                if (elementsToProcess.Count < 2)
                {
                    TaskDialog.Show("Warning", "Cần chọn ít nhất 2 đối tượng để so sánh.");
                    return Result.Succeeded;
                }
            }
            else
            {
                elementsToProcess = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .ToElements();
            }

            // 2. PROCESSING: Tìm trùng lặp
            List<DuplicatePair> duplicates = FindDuplicates(doc, elementsToProcess);

            if (duplicates.Count == 0)
            {
                TaskDialog.Show("Kết quả", "Không tìm thấy đối tượng trùng lặp (Giao nhau > 60%).");
                return Result.Succeeded;
            }

            // 3. OUTPUT: Hiển thị Form
            CheckRequestHandler handler = new CheckRequestHandler();
            ExternalEvent exEvent = ExternalEvent.Create(handler);

            _openedForm = new DuplicateResultForm(uiDoc, duplicates, exEvent, handler);
            _openedForm.Show();

            return Result.Succeeded;
        }

        private List<DuplicatePair> FindDuplicates(Document doc, IList<Element> elements)
        {
            var pairs = new List<DuplicatePair>();
            var checkedIds = new HashSet<ElementId>();

            // --- FIX QUAN TRỌNG: GROUP BY CATEGORY ---
            // Code cũ group theo Symbol.Id (Type), dẫn đến việc khác Type sẽ không so sánh.
            // Code mới chỉ group theo Category (VD: Cột so với Cột, bất kể Type gì).
            var grouped = elements.GroupBy(e =>
            {
                if (e.Category == null) return -1;
                return e.Category.Id.IntegerValue;
            });

            // Cấu hình lấy Geometry chính xác (Fine)
            Options opt = new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = true };

            foreach (var group in grouped)
            {
                var list = group.ToList();
                if (list.Count < 2) continue;

                for (int i = 0; i < list.Count; i++)
                {
                    Element e1 = list[i];
                    if (checkedIds.Contains(e1.Id)) continue;

                    // Lấy Solid E1
                    Solid solid1 = GetSolidFromElement(e1, opt);
                    if (solid1 == null || solid1.Volume <= 0.0001) continue;

                    BoundingBoxXYZ bb1 = e1.get_BoundingBox(null);

                    for (int j = i + 1; j < list.Count; j++)
                    {
                        Element e2 = list[j];
                        if (checkedIds.Contains(e2.Id)) continue;

                        // Bước 1: Check BoundingBox (Nhanh)
                        BoundingBoxXYZ bb2 = e2.get_BoundingBox(null);
                        if (bb1 == null || bb2 == null || !BoundingBoxIntersects(bb1, bb2)) continue;

                        // Bước 2: Check Solid (Chính xác)
                        Solid solid2 = GetSolidFromElement(e2, opt);
                        if (solid2 == null || solid2.Volume <= 0.0001) continue;

                        // Ngưỡng trùng lặp: 60% thể tích
                        if (IsDuplicateBySolidOverlap(solid1, solid2, 0.60))
                        {
                            pairs.Add(new DuplicatePair { KeepElement = e1, DeleteElement = e2 });
                            checkedIds.Add(e2.Id);
                        }
                    }
                }
            }
            return pairs;
        }

        private bool BoundingBoxIntersects(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
        {
            return (bb1.Max.X >= bb2.Min.X && bb1.Min.X <= bb2.Max.X) &&
                   (bb1.Max.Y >= bb2.Min.Y && bb1.Min.Y <= bb2.Max.Y) &&
                   (bb1.Max.Z >= bb2.Min.Z && bb1.Min.Z <= bb2.Max.Z);
        }

        private Solid GetSolidFromElement(Element e, Options opt)
        {
            GeometryElement geoElem = e.get_Geometry(opt);
            if (geoElem == null) return null;
            return GetSolidFromGeometryElement(geoElem);
        }

        private Solid GetSolidFromGeometryElement(GeometryElement geoElem)
        {
            foreach (GeometryObject obj in geoElem)
            {
                if (obj is Solid solid && solid.Volume > 0.0001)
                {
                    return solid;
                }
                else if (obj is GeometryInstance geomInst)
                {
                    // Đệ quy để lấy Solid bên trong Group/Family Instance
                    // Quan trọng: Dùng GetInstanceGeometry() để lấy toạ độ thực tế
                    Solid instSolid = GetSolidFromGeometryElement(geomInst.GetInstanceGeometry());
                    if (instSolid != null && instSolid.Volume > 0.0001) return instSolid;
                }
            }
            return null;
        }

        private bool IsDuplicateBySolidOverlap(Solid s1, Solid s2, double percentThreshold)
        {
            try
            {
                // Tìm phần giao nhau
                Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(s1, s2, BooleanOperationsType.Intersect);

                if (intersection == null || intersection.Volume <= 0.00001) return false;

                // Logic: (Thể tích giao) / (Thể tích vật nhỏ hơn)
                double minVolume = Math.Min(s1.Volume, s2.Volume);
                if (minVolume <= 0.00001) return false;

                double ratio = intersection.Volume / minVolume;

                return ratio >= percentThreshold;
            }
            catch
            {
                return false;
            }
        }
    }

    // --- DATA MODEL ---
    public class DuplicatePair
    {
        public Element KeepElement { get; set; }
        public Element DeleteElement { get; set; }
    }

    // --- EXTERNAL EVENT HANDLER ---
    public enum RequestId { None, SectionBox, Delete }

    public class CheckRequestHandler : IExternalEventHandler
    {
        public RequestId Request { get; set; } = RequestId.None;
        public List<DuplicatePair> CurrentData { get; set; }
        public int SelectedIndex { get; set; } = -1;
        public List<ElementId> IdsPendingDelete { get; set; } = new List<ElementId>();

        public void Execute(UIApplication app)
        {
            try
            {
                switch (Request)
                {
                    case RequestId.SectionBox: DoSectionBox(app); break;
                    case RequestId.Delete: DoDelete(app); break;
                }
            }
            catch (Exception ex) { TaskDialog.Show("Error", ex.Message); }
        }

        private void DoSectionBox(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            Document doc = uiDoc.Document;
            if (SelectedIndex < 0 || SelectedIndex >= CurrentData.Count) return;
            var pair = CurrentData[SelectedIndex];

            View3D targetView = GetAndActivate3DView(uiDoc);
            if (targetView == null)
            {
                TaskDialog.Show("Error", "Không tìm thấy View 3D nào.");
                return;
            }

            using (Transaction t = new Transaction(doc, "Auto Section Box"))
            {
                t.Start();
                if (!targetView.IsSectionBoxActive) targetView.IsSectionBoxActive = true;

                BoundingBoxXYZ box1 = pair.KeepElement.get_BoundingBox(null);
                BoundingBoxXYZ box2 = pair.DeleteElement.get_BoundingBox(null);

                if (box1 != null && box2 != null)
                {
                    // Tạo hộp bao gộp cả 2 đối tượng
                    XYZ minPt = new XYZ(Math.Min(box1.Min.X, box2.Min.X), Math.Min(box1.Min.Y, box2.Min.Y), Math.Min(box1.Min.Z, box2.Min.Z));
                    XYZ maxPt = new XYZ(Math.Max(box1.Max.X, box2.Max.X), Math.Max(box1.Max.Y, box2.Max.Y), Math.Max(box1.Max.Z, box2.Max.Z));

                    // Thêm padding 2 feet (~600mm)
                    double padding = 2.0;
                    BoundingBoxXYZ newBox = new BoundingBoxXYZ();
                    newBox.Min = minPt - new XYZ(padding, padding, padding);
                    newBox.Max = maxPt + new XYZ(padding, padding, padding);

                    targetView.SetSectionBox(newBox);
                }
                t.Commit();
            }

            // Select 2 đối tượng để người dùng dễ thấy
            uiDoc.Selection.SetElementIds(new List<ElementId> { pair.KeepElement.Id, pair.DeleteElement.Id });

            // Zoom tới đối tượng
            var uiview = uiDoc.GetOpenUIViews().FirstOrDefault(uv => uv.ViewId == targetView.Id);
            if (uiview != null) uiview.ZoomToFit();
        }

        private void DoDelete(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            using (Transaction t = new Transaction(doc, "Delete Duplicates"))
            {
                t.Start();
                int successCount = 0;
                if (IdsPendingDelete != null && IdsPendingDelete.Count > 0)
                {
                    // Validate ID trước khi xóa để tránh crash
                    List<ElementId> validIds = new List<ElementId>();
                    foreach (var id in IdsPendingDelete)
                    {
                        if (doc.GetElement(id) != null) validIds.Add(id);
                    }
                    if (validIds.Count > 0)
                    {
                        doc.Delete(validIds);
                        successCount = validIds.Count;
                    }
                }
                t.Commit();
                TaskDialog.Show("Result", $"Đã xóa {successCount} đối tượng.");
            }
            // Đóng form sau khi xóa
            if (DuplicateElementsDetectorCmd._openedForm != null) DuplicateElementsDetectorCmd._openedForm.Close();
        }

        private View3D GetAndActivate3DView(UIDocument uiDoc)
        {
            Document doc = uiDoc.Document;
            if (doc.ActiveView is View3D active3D && !active3D.IsTemplate && !active3D.IsAssemblyView) return active3D;

            var collector = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().Where(v => !v.IsTemplate && !v.IsAssemblyView);

            // Ưu tiên view có tên "{3D}"
            View3D targetView = collector.FirstOrDefault(v => v.Name.Contains("{3D}"));
            if (targetView == null) targetView = collector.FirstOrDefault();

            if (targetView != null) uiDoc.ActiveView = targetView;
            return targetView;
        }

        public string GetName() => "Duplicate Check Handler";
    }

    // --- USER INTERFACE (FORM) ---
    public class DuplicateResultForm : Form
    {
        private UIDocument _uiDoc;
        private List<DuplicatePair> _data;
        private ExternalEvent _exEvent;
        private CheckRequestHandler _handler;

        private DataGridView dgv;
        private Button btnDeleteSelected, btnDeleteAll, btnClose;
        private CheckBox chkSelectAll;

        public DuplicateResultForm(UIDocument uiDoc, List<DuplicatePair> data, ExternalEvent exEvent, CheckRequestHandler handler)
        {
            _uiDoc = uiDoc; _data = data; _exEvent = exEvent; _handler = handler;
            InitializeComponent(); LoadData();
        }

        private void InitializeComponent()
        {
            this.Text = "Duplicate Detector Results";
            this.Size = new Size(600, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;

            var mainTable = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, Padding = new Padding(10) };
            mainTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // Top Panel
            var pnlTop = new Panel { Height = 30, Dock = DockStyle.Fill };
            chkSelectAll = new CheckBox { Text = "Select All", Location = new Point(0, 5), AutoSize = true };
            chkSelectAll.CheckedChanged += (s, e) => { foreach (DataGridViewRow r in dgv.Rows) r.Cells[0].Value = chkSelectAll.Checked; };
            pnlTop.Controls.Add(chkSelectAll);

            // Grid
            dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, BackgroundColor = Color.White };
            dgv.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "✔", Width = 30 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Keep ID", Width = 80, ReadOnly = true });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Delete ID", Width = 80, ReadOnly = true });
            dgv.Columns.Add(new DataGridViewButtonColumn { HeaderText = "Action", Text = "Show", UseColumnTextForButtonValue = true });

            dgv.CellContentClick += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == 3) // Nút Show
                {
                    _handler.Request = RequestId.SectionBox;
                    _handler.CurrentData = _data;
                    _handler.SelectedIndex = e.RowIndex;
                    _exEvent.Raise();
                }
            };

            // Bottom Panel
            var pnlBot = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Height = 40 };
            btnClose = new Button { Text = "Close", Width = 80 };
            btnClose.Click += (s, e) => this.Close();

            btnDeleteAll = new Button { Text = "Delete All", Width = 100, BackColor = Color.LightCoral };
            btnDeleteAll.Click += (s, e) =>
            {
                if (MessageBox.Show($"Delete ALL {_data.Count} duplicates?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    _handler.Request = RequestId.Delete;
                    _handler.IdsPendingDelete = _data.Select(x => x.DeleteElement.Id).ToList();
                    _exEvent.Raise();
                }
            };

            btnDeleteSelected = new Button { Text = "Delete Selected", Width = 120 };
            btnDeleteSelected.Click += (s, e) =>
            {
                List<ElementId> ids = new List<ElementId>();
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (Convert.ToBoolean(row.Cells[0].Value))
                        ids.Add(new ElementId(int.Parse(row.Cells[2].Value.ToString())));
                }
                if (ids.Count > 0)
                {
                    _handler.Request = RequestId.Delete;
                    _handler.IdsPendingDelete = ids;
                    _exEvent.Raise();
                }
            };

            pnlBot.Controls.Add(btnClose);
            pnlBot.Controls.Add(btnDeleteAll);
            pnlBot.Controls.Add(btnDeleteSelected);

            mainTable.Controls.Add(pnlTop, 0, 0);
            mainTable.Controls.Add(dgv, 0, 1);
            mainTable.Controls.Add(pnlBot, 0, 2);

            this.Controls.Add(mainTable);
        }

        private void LoadData()
        {
            foreach (var p in _data)
            {
                dgv.Rows.Add(false, p.KeepElement.Id.IntegerValue, p.DeleteElement.Id.IntegerValue);
            }
        }
    }
}