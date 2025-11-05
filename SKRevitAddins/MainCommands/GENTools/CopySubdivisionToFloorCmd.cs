using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ComboBox = System.Windows.Forms.ComboBox;
using Control = System.Windows.Forms.Control;
using Form = System.Windows.Forms.Form;
using Level = Autodesk.Revit.DB.Level;
using Floor = Autodesk.Revit.DB.Floor;

namespace SKRevitAddins.GENTools
{
    [Transaction(TransactionMode.Manual)]
    public class CopySubdivisionToFloorCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string msg, ElementSet set)
        {
            UIDocument uidoc = c.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // --- Chọn Toposolid/Sub-Division ---
                Reference picked = uidoc.Selection.PickObject(ObjectType.Element, "Chọn Toposolid hoặc Sub-Division");
                if (picked == null) return Result.Cancelled;

                Toposolid topo = doc.GetElement(picked) as Toposolid;
                if (topo == null)
                {
                    TaskDialog.Show("Lỗi", "Phần tử chọn không phải Toposolid/Sub-Division.");
                    return Result.Cancelled;
                }

                // --- Tính cao độ trung bình ---
                double avgZ = GetAverageElevation(topo);

                // --- Lấy danh sách loại sàn ---
                var floorTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(FloorType))
                    .Cast<FloorType>()
                    .OrderBy(ft => ft.Name)
                    .ToList();

                if (floorTypes.Count == 0)
                {
                    TaskDialog.Show("Lỗi", "Không tìm thấy FloorType trong dự án.");
                    return Result.Failed;
                }

                // --- Mở form chọn loại sàn ---
                FloorType selectedType;
                using (var form = new SelectFloorTypeForm(floorTypes))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;
                    selectedType = form.SelectedType;
                }

                // --- Tìm Level gần nhất ---
                Level baseLevel = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(lv => Math.Abs(lv.Elevation - avgZ))
                    .FirstOrDefault();

                if (baseLevel == null)
                {
                    TaskDialog.Show("Lỗi", "Không tìm thấy Level hợp lệ.");
                    return Result.Failed;
                }

                // --- Lấy biên dạng từ Sketch hoặc Geometry ---
                IList<CurveLoop> loops = GetBoundaryLoops(doc, topo);
                if (loops == null || loops.Count == 0)
                {
                    TaskDialog.Show("Lỗi", "Không lấy được biên dạng Toposolid.");
                    return Result.Failed;
                }

                using (Transaction t = new Transaction(doc, "Copy Toposolid to Floor"))
                {
                    t.Start();

                    // --- Làm phẳng TẤT CẢ loops ---
                    List<CurveLoop> flatLoops = new List<CurveLoop>();
                    const double tolerance = 1e-6;

                    foreach (CurveLoop originalLoop in loops)
                    {
                        List<Curve> flatCurves = new List<Curve>();

                        foreach (Curve curve in originalLoop)
                        {
                            XYZ p1 = new XYZ(curve.GetEndPoint(0).X, curve.GetEndPoint(0).Y, avgZ);
                            XYZ p2 = new XYZ(curve.GetEndPoint(1).X, curve.GetEndPoint(1).Y, avgZ);

                            if (p1.DistanceTo(p2) < tolerance)
                                continue;

                            if (curve is Line)
                            {
                                flatCurves.Add(Line.CreateBound(p1, p2));
                            }
                            else if (curve is Arc arc)
                            {
                                XYZ midOriginal = arc.Evaluate(0.5, true);
                                XYZ pMid = new XYZ(midOriginal.X, midOriginal.Y, avgZ);

                                try
                                {
                                    flatCurves.Add(Arc.Create(p1, p2, pMid));
                                }
                                catch
                                {
                                    flatCurves.Add(Line.CreateBound(p1, p2));
                                }
                            }
                            else
                            {
                                IList<XYZ> tessellated = curve.Tessellate();
                                for (int i = 0; i < tessellated.Count - 1; i++)
                                {
                                    XYZ pt1 = new XYZ(tessellated[i].X, tessellated[i].Y, avgZ);
                                    XYZ pt2 = new XYZ(tessellated[i + 1].X, tessellated[i + 1].Y, avgZ);

                                    if (pt1.DistanceTo(pt2) >= tolerance)
                                        flatCurves.Add(Line.CreateBound(pt1, pt2));
                                }
                            }
                        }

                        if (flatCurves.Count >= 3)
                        {
                            try
                            {
                                CurveLoop flatLoop = CurveLoop.Create(flatCurves);
                                flatLoops.Add(flatLoop);
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("Cảnh báo", $"Bỏ qua loop không hợp lệ: {ex.Message}");
                            }
                        }
                    }

                    if (flatLoops.Count == 0)
                    {
                        TaskDialog.Show("Lỗi", "Không có loop hợp lệ để tạo sàn.");
                        t.RollBack();
                        return Result.Failed;
                    }

                    // --- Tạo sàn ---
                    Floor floor;
                    try
                    {
                        floor = Floor.Create(doc, flatLoops, selectedType.Id, baseLevel.Id);
                        floor.Name = $"Floor_from_{topo.Name}";
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Lỗi tạo sàn", ex.Message);
                        t.RollBack();
                        return Result.Failed;
                    }

                    // --- Copy vật liệu ---
                    ICollection<ElementId> mats = topo.GetMaterialIds(false);
                    if (mats != null && mats.Count > 0)
                    {
                        ElementId matId = mats.First();
                        Parameter matParam = floor.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM);
                        if (matParam != null && !matParam.IsReadOnly)
                        {
                            try { matParam.Set(matId); } catch { }
                        }
                    }

                    // --- Đặt cao độ ---
                    Parameter heightParam = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                    if (heightParam != null && !heightParam.IsReadOnly)
                        heightParam.Set(avgZ - baseLevel.Elevation);

                    t.Commit();

                    string loopInfo = flatLoops.Count == 1
                        ? "1 loop (outer boundary)"
                        : $"{flatLoops.Count} loops (outer + {flatLoops.Count - 1} voids)";

                    TaskDialog.Show("Hoàn tất",
                        $"✅ Đã tạo sàn '{floor.Name}' với {loopInfo}\n" +
                        $"Từ Toposolid '{topo.Name}' tại Level '{baseLevel.Name}'.");
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                return Result.Failed;
            }
        }

        // --- Lấy boundary từ Sketch hoặc Geometry ---
        private IList<CurveLoop> GetBoundaryLoops(Document doc, Toposolid topo)
        {
            // PHƯƠNG PHÁP 1: Lấy từ Sketch (chính xác nhất)
            Sketch sketch = doc.GetElement(topo.SketchId) as Sketch;
            if (sketch != null)
            {
                var profile = sketch.Profile;
                if (profile != null && profile.Size > 0)
                {
                    List<CurveLoop> loops = new List<CurveLoop>();
                    foreach (CurveArray curveArray in profile)
                    {
                        List<Curve> curves = new List<Curve>();
                        foreach (Curve c in curveArray)
                            curves.Add(c);

                        if (curves.Count >= 3)
                        {
                            try
                            {
                                loops.Add(CurveLoop.Create(curves));
                            }
                            catch { }
                        }
                    }

                    if (loops.Count > 0)
                        return loops;
                }
            }

            // PHƯƠNG PHÁP 2: Phân tích Solid edges (fallback)
            return GetBoundaryFromGeometry(topo);
        }

        // --- Phương pháp dự phòng: Phân tích Solid ---
        private IList<CurveLoop> GetBoundaryFromGeometry(Toposolid topo)
        {
            Options opt = new Options()
            {
                ComputeReferences = true,
                DetailLevel = ViewDetailLevel.Fine
            };
            GeometryElement geo = topo.get_Geometry(opt);

            // Tìm tất cả các mặt hướng lên
            List<Face> topFaces = new List<Face>();
            foreach (GeometryObject obj in geo)
            {
                if (obj is Solid s && s.Faces.Size > 0)
                {
                    foreach (Face f in s.Faces)
                    {
                        XYZ normal = f.ComputeNormal(new UV(0.5, 0.5));
                        if (normal.Z > 0.7) // Mặt hướng lên
                        {
                            topFaces.Add(f);
                        }
                    }
                }
            }

            if (topFaces.Count == 0)
                return new List<CurveLoop>();

            // Lấy mặt có diện tích lớn nhất
            Face mainFace = topFaces.OrderByDescending(f => f.Area).First();

            // Lấy tất cả edges của mặt này
            var loops = mainFace.GetEdgesAsCurveLoops();

            // Nếu có nhiều face liền kề, thử merge edges
            if (topFaces.Count > 1)
            {
                List<CurveLoop> allLoops = new List<CurveLoop>();
                foreach (var face in topFaces)
                {
                    var faceLoops = face.GetEdgesAsCurveLoops();
                    foreach (var loop in faceLoops)
                        allLoops.Add(loop);
                }

                // Loại bỏ loops trùng lặp dựa trên bounding box
                var uniqueLoops = RemoveDuplicateLoops(allLoops);
                if (uniqueLoops.Count > loops.Count)
                    return uniqueLoops;
            }

            return loops;
        }

        // --- Loại bỏ loops trùng lặp ---
        private List<CurveLoop> RemoveDuplicateLoops(List<CurveLoop> loops)
        {
            List<CurveLoop> unique = new List<CurveLoop>();
            const double tolerance = 0.01; // 1cm

            foreach (var loop in loops)
            {
                bool isDuplicate = false;

                // Tính bounding box thủ công
                List<XYZ> points = new List<XYZ>();
                foreach (Curve c in loop)
                {
                    points.Add(c.GetEndPoint(0));
                    points.Add(c.GetEndPoint(1));
                }

                double minX = points.Min(p => p.X);
                double minY = points.Min(p => p.Y);
                double maxX = points.Max(p => p.X);
                double maxY = points.Max(p => p.Y);

                foreach (var existingLoop in unique)
                {
                    // Tính bbox của existing loop
                    List<XYZ> existingPoints = new List<XYZ>();
                    foreach (Curve c in existingLoop)
                    {
                        existingPoints.Add(c.GetEndPoint(0));
                        existingPoints.Add(c.GetEndPoint(1));
                    }

                    double exMinX = existingPoints.Min(p => p.X);
                    double exMinY = existingPoints.Min(p => p.Y);
                    double exMaxX = existingPoints.Max(p => p.X);
                    double exMaxY = existingPoints.Max(p => p.Y);

                    // So sánh bounding box
                    if (Math.Abs(minX - exMinX) < tolerance &&
                        Math.Abs(minY - exMinY) < tolerance &&
                        Math.Abs(maxX - exMaxX) < tolerance &&
                        Math.Abs(maxY - exMaxY) < tolerance)
                    {
                        isDuplicate = true;
                        break;
                    }
                }

                if (!isDuplicate)
                    unique.Add(loop);
            }

            return unique;
        }

        // --- Tính cao độ trung bình ---
        private double GetAverageElevation(Toposolid topo)
        {
            Options opt = new Options() { DetailLevel = ViewDetailLevel.Fine };
            GeometryElement geo = topo.get_Geometry(opt);
            List<double> zs = new List<double>();

            foreach (GeometryObject obj in geo)
            {
                if (obj is Solid s)
                {
                    foreach (Face f in s.Faces)
                    {
                        XYZ normal = f.ComputeNormal(new UV(0.5, 0.5));
                        if (normal.Z > 0.5)
                        {
                            Mesh m = f.Triangulate();
                            foreach (XYZ p in m.Vertices)
                                zs.Add(p.Z);
                        }
                    }
                }
            }
            return zs.Count > 0 ? zs.Average() : 0.0;
        }
    }

    // --- Form chọn loại sàn ---
    public class SelectFloorTypeForm : Form
    {
        private ComboBox cbType;
        private Button btnOK, btnCancel;
        private readonly List<FloorType> _types;
        public FloorType SelectedType { get; private set; }

        public SelectFloorTypeForm(List<FloorType> types)
        {
            _types = types;

            Text = "Chọn loại sàn";
            Width = 380;
            Height = 150;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            Label lbl = new Label()
            {
                Text = "Loại sàn:",
                Left = 20,
                Top = 25,
                Width = 70
            };

            cbType = new ComboBox()
            {
                Left = 100,
                Top = 20,
                Width = 240,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            foreach (var t in _types)
                cbType.Items.Add(t.Name);

            if (cbType.Items.Count > 0)
                cbType.SelectedIndex = 0;

            btnOK = new Button()
            {
                Text = "OK",
                Left = 180,
                Width = 70,
                Top = 70,
                DialogResult = DialogResult.OK
            };
            btnCancel = new Button()
            {
                Text = "Hủy",
                Left = 270,
                Width = 70,
                Top = 70,
                DialogResult = DialogResult.Cancel
            };

            Controls.AddRange(new Control[] { lbl, cbType, btnOK, btnCancel });
            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                string name = cbType.SelectedItem?.ToString();
                SelectedType = _types.FirstOrDefault(ft =>
                    ft.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            }
            base.OnFormClosing(e);
        }
    }
}