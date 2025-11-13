using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;
using Color = System.Drawing.Color;
using ComboBox = System.Windows.Forms.ComboBox;
using Control = System.Windows.Forms.Control;
using Form = System.Windows.Forms.Form;
using Point = System.Drawing.Point;
using TextBox = System.Windows.Forms.TextBox;
using UnitUtils = Autodesk.Revit.DB.UnitUtils;

namespace SKRevitAddins.PointCloudAddins
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CheckAndMapPointCloudCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string msg, ElementSet set)
        {
            var uiDoc = c.Application.ActiveUIDocument;
            var doc = uiDoc?.Document;
            if (doc == null)
            {
                msg = "Không tìm thấy tài liệu Revit hiện hành.";
                return Result.Failed;
            }

            try
            {
                var pcInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(PointCloudInstance))
                    .Cast<PointCloudInstance>()
                    .ToList();

                if (pcInstances.Count < 2)
                {
                    TaskDialog.Show("PointCloud", "Cần tối thiểu 2 Point Cloud đã được link sẵn trong dự án.");
                    return Result.Succeeded;
                }

                var infos = CollectPointCloudInfos(pcInstances)
                    .OrderBy(i => i.Name, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

                using (var form = new MapExistingPointCloudsForm(infos))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;

                    ElementId subjectId = form.SubjectPointCloudId;
                    ElementId referenceId = form.ReferencePointCloudId;

                    if (subjectId == ElementId.InvalidElementId || referenceId == ElementId.InvalidElementId)
                    {
                        TaskDialog.Show("Lỗi", "Vui lòng chọn đầy đủ 2 Point Cloud.");
                        return Result.Failed;
                    }
                    if (subjectId == referenceId)
                    {
                        TaskDialog.Show("Lỗi", "Bên trái (subject) và bên phải (reference) phải khác nhau.");
                        return Result.Failed;
                    }

                    var subject = pcInstances.FirstOrDefault(p => p.Id == subjectId);
                    var referencePc = pcInstances.FirstOrDefault(p => p.Id == referenceId);
                    if (subject == null || referencePc == null)
                    {
                        TaskDialog.Show("Lỗi", "Không tìm thấy Point Cloud đã chọn.");
                        return Result.Failed;
                    }

                    using (var t = new Transaction(doc, "Map PointCloud (subject → reference)"))
                    {
                        t.Start();

                        MapPointCloud(doc, subject, referencePc,
                            form.OffsetXmm, form.OffsetYmm, form.OffsetZmm,
                            form.RotationAddDeg);

                        t.Commit();
                    }

                    TaskDialog.Show("Shinken Tools - Hoàn tất",
                        $"Đã mapping:\n" +
                        $"- Bên trái (đã chỉnh): {subject.Name}\n" +
                        $"- Bên phải (chuẩn): {referencePc.Name}\n\n" +
                        $"Offset thêm: ({form.OffsetXmm}, {form.OffsetYmm}, {form.OffsetZmm}) mm\n" +
                        $"Góc xoay thêm: {form.RotationAddDeg}°");

                    return Result.Succeeded;
                }
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                TaskDialog.Show("Shinken Tools - Lỗi", ex.Message);
                return Result.Failed;
            }
        }

        private static IEnumerable<PcInfo> CollectPointCloudInfos(IEnumerable<PointCloudInstance> pcs)
        {
            foreach (var pc in pcs)
            {
                Transform tf = pc.GetTotalTransform();
                XYZ o = tf.Origin;

                double x = UnitUtils.ConvertFromInternalUnits(o.X, UnitTypeId.Millimeters);
                double y = UnitUtils.ConvertFromInternalUnits(o.Y, UnitTypeId.Millimeters);
                double z = UnitUtils.ConvertFromInternalUnits(o.Z, UnitTypeId.Millimeters);
                double rotDeg = Math.Atan2(tf.BasisX.Y, tf.BasisX.X) * 180.0 / Math.PI;

                yield return new PcInfo
                {
                    Id = pc.Id,
                    Name = pc.Name,
                    Xmm = x,
                    Ymm = y,
                    Zmm = z,
                    RotDeg = rotDeg,
                    TotalTransform = tf
                };
            }
        }

        private static void MapPointCloud(
            Document doc,
            PointCloudInstance subject,
            PointCloudInstance referencePc,
            double offsetXmm,
            double offsetYmm,
            double offsetZmm,
            double rotationAddDeg)
        {
            Transform T_subject_current = subject.GetTotalTransform();
            Transform T_reference = referencePc.GetTotalTransform();

            double dx = UnitUtils.ConvertToInternalUnits(offsetXmm, UnitTypeId.Millimeters);
            double dy = UnitUtils.ConvertToInternalUnits(offsetYmm, UnitTypeId.Millimeters);
            double dz = UnitUtils.ConvertToInternalUnits(offsetZmm, UnitTypeId.Millimeters);
            double rotRad = rotationAddDeg * Math.PI / 180.0;

            Transform T_move = Transform.CreateTranslation(new XYZ(dx, dy, dz));
            Transform T_rot = Transform.CreateRotation(XYZ.BasisZ, rotRad);
            Transform T_subject_new = T_reference.Multiply(T_rot).Multiply(T_move);

            Transform Delta = T_subject_current.Inverse.Multiply(T_subject_new);

            var pcElem = doc.GetElement(subject.Id);
            bool wasPinned = pcElem.Pinned;
            if (wasPinned) pcElem.Pinned = false;

            XYZ originSubject = T_subject_current.Origin;
            double deltaRotZ = Math.Atan2(Delta.BasisX.Y, Delta.BasisX.X);
            Line axis = Line.CreateBound(originSubject, originSubject + XYZ.BasisZ);
            ElementTransformUtils.RotateElement(doc, subject.Id, axis, deltaRotZ);

            XYZ tr = Delta.Origin;
            ElementTransformUtils.MoveElement(doc, subject.Id, tr);

            if (wasPinned && pcElem != null) pcElem.Pinned = true;
        }

        private class PcInfo
        {
            public ElementId Id { get; set; }
            public string Name { get; set; }
            public double Xmm { get; set; }
            public double Ymm { get; set; }
            public double Zmm { get; set; }
            public double RotDeg { get; set; }
            public Transform TotalTransform { get; set; }
            public override string ToString() => Name;
        }

        private class MapExistingPointCloudsForm : Form
        {
            private readonly List<PcInfo> _pcs;
            private ComboBox cmbSubject;
            private ComboBox cmbReference;
            private Label lblSubX, lblSubY, lblSubZ, lblSubRot;
            private Label lblRefX, lblRefY, lblRefZ, lblRefRot;
            private TextBox txtOffX, txtOffY, txtOffZ, txtRot;
            private Button btnSwap, btnOK, btnCancel;

            public ElementId SubjectPointCloudId =>
                (cmbSubject.SelectedItem as PcInfo)?.Id ?? ElementId.InvalidElementId;

            public ElementId ReferencePointCloudId =>
                (cmbReference.SelectedItem as PcInfo)?.Id ?? ElementId.InvalidElementId;

            public double OffsetXmm => ParseDouble(txtOffX.Text);
            public double OffsetYmm => ParseDouble(txtOffY.Text);
            public double OffsetZmm => ParseDouble(txtOffZ.Text);
            public double RotationAddDeg => ParseDouble(txtRot.Text);

            public MapExistingPointCloudsForm(List<PcInfo> pcs)
            {
                _pcs = pcs ?? new List<PcInfo>();

                Text = "Shinken Group® - Map 2 Point Cloud đã link";
                Size = new Size(900, 460);
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                BackColor = Color.WhiteSmoke;

                var logo = new PictureBox
                {
                    Left = 20,
                    Top = 10,
                    Width = 120,
                    Height = 40,
                    SizeMode = PictureBoxSizeMode.Zoom
                };
                AddLogoIfFound(logo);
                Controls.Add(logo);

                var lblTitle = new Label
                {
                    Text = "Chọn 2 Point Cloud đã link sẵn: BÊN TRÁI sẽ chỉnh để khớp BÊN PHẢI",
                    Location = new Point(160, 18),
                    AutoSize = true,
                    Font = new Font("Segoe UI", 10, FontStyle.Bold)
                };

                var lblSubject = new Label { Text = "Bên trái (sẽ chỉnh):", Left = 20, Top = 70, Width = 150 };
                cmbSubject = new ComboBox { Left = 170, Top = 66, Width = 500, DropDownStyle = ComboBoxStyle.DropDownList };
                cmbSubject.DataSource = _pcs.ToList();
                cmbSubject.DisplayMember = "Name";
                cmbSubject.SelectedIndexChanged += (s, e) => UpdateInfoPanels();

                var lblReference = new Label { Text = "Bên phải (chuẩn):", Left = 20, Top = 100, Width = 150 };
                cmbReference = new ComboBox { Left = 170, Top = 96, Width = 500, DropDownStyle = ComboBoxStyle.DropDownList };
                cmbReference.DataSource = _pcs.ToList();
                cmbReference.DisplayMember = "Name";
                cmbReference.SelectedIndexChanged += (s, e) => UpdateInfoPanels();

                btnSwap = new Button
                {
                    Text = "Hoán đổi trái ↔ phải",
                    Left = 170,
                    Top = 128,
                    Width = 200,
                    Height = 28,
                    BackColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnSwap.FlatAppearance.BorderColor = Color.Silver;
                btnSwap.Click += (s, e) => SwapSelections();

                var grpLeft = new GroupBox
                {
                    Text = "Bên trái (sẽ chỉnh)",
                    Left = 20,
                    Top = 170,
                    Width = 340,
                    Height = 140
                };
                lblSubX = new Label { Left = 15, Top = 25, Width = 300 };
                lblSubY = new Label { Left = 15, Top = 50, Width = 300 };
                lblSubZ = new Label { Left = 15, Top = 75, Width = 300 };
                lblSubRot = new Label { Left = 15, Top = 100, Width = 300 };
                grpLeft.Controls.AddRange(new Control[] { lblSubX, lblSubY, lblSubZ, lblSubRot });

                var grpRight = new GroupBox
                {
                    Text = "Bên phải (chuẩn)",
                    Left = 370,
                    Top = 170,
                    Width = 340,
                    Height = 140
                };
                lblRefX = new Label { Left = 15, Top = 25, Width = 300 };
                lblRefY = new Label { Left = 15, Top = 50, Width = 300 };
                lblRefZ = new Label { Left = 15, Top = 75, Width = 300 };
                lblRefRot = new Label { Left = 15, Top = 100, Width = 300 };
                grpRight.Controls.AddRange(new Control[] { lblRefX, lblRefY, lblRefZ, lblRefRot });

                var sep = new Label { BorderStyle = BorderStyle.Fixed3D, Left = 20, Top = 330, Width = 690, Height = 2 };

                var lblOff = new Label { Text = "Offset thêm (mm):", Left = 20, Top = 345, Width = 130 };
                txtOffX = new TextBox { Left = 160, Top = 342, Width = 80, Text = "0" };
                txtOffY = new TextBox { Left = 245, Top = 342, Width = 80, Text = "0" };
                txtOffZ = new TextBox { Left = 330, Top = 342, Width = 80, Text = "0" };

                var lblRot = new Label { Text = "Góc xoay thêm (°):", Left = 20, Top = 375, Width = 130 };
                txtRot = new TextBox { Left = 160, Top = 372, Width = 80, Text = "0" };

                btnOK = new Button
                {
                    Text = "Áp dụng",
                    Left = 510,
                    Top = 370,
                    Width = 90,
                    Height = 32,
                    DialogResult = DialogResult.OK,
                    BackColor = Color.FromArgb(0, 120, 215),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnOK.FlatAppearance.BorderSize = 0;

                btnCancel = new Button
                {
                    Text = "Hủy",
                    Left = 610,
                    Top = 370,
                    Width = 90,
                    Height = 32,
                    DialogResult = DialogResult.Cancel,
                    BackColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnCancel.FlatAppearance.BorderColor = Color.Silver;

                Controls.AddRange(new Control[]
                {
                    lblTitle,
                    lblSubject, cmbSubject,
                    lblReference, cmbReference,
                    btnSwap,
                    grpLeft, grpRight,
                    sep,
                    lblOff, txtOffX, txtOffY, txtOffZ,
                    lblRot, txtRot,
                    btnOK, btnCancel
                });

                AcceptButton = btnOK;
                CancelButton = btnCancel;

                Load += (s, e) =>
                {
                    if (_pcs.Count >= 2)
                    {
                        cmbSubject.SelectedIndex = 0;
                        cmbReference.SelectedIndex = Math.Min(1, _pcs.Count - 1);
                    }
                    UpdateInfoPanels();
                };
            }

            private void SwapSelections()
            {
                int iSub = cmbSubject.SelectedIndex;
                int iRef = cmbReference.SelectedIndex;

                if (iSub < 0 && iRef < 0) return;

                cmbSubject.SelectedIndex = iRef >= 0 ? iRef : 0;
                cmbReference.SelectedIndex = iSub >= 0 ? iSub : 0;

                UpdateInfoPanels();
            }

            private void UpdateInfoPanels()
            {
                var sub = cmbSubject.SelectedItem as PcInfo;
                var refe = cmbReference.SelectedItem as PcInfo;

                if (sub != null)
                {
                    lblSubX.Text = $"X: {sub.Xmm:F1} mm";
                    lblSubY.Text = $"Y: {sub.Ymm:F1} mm";
                    lblSubZ.Text = $"Z: {sub.Zmm:F1} mm";
                    lblSubRot.Text = $"Góc xoay: {sub.RotDeg:F2}°";
                }
                else
                {
                    lblSubX.Text = "X: —";
                    lblSubY.Text = "Y: —";
                    lblSubZ.Text = "Z: —";
                    lblSubRot.Text = "Góc xoay: —";
                }

                if (refe != null)
                {
                    lblRefX.Text = $"X: {refe.Xmm:F1} mm";
                    lblRefY.Text = $"Y: {refe.Ymm:F1} mm";
                    lblRefZ.Text = $"Z: {refe.Zmm:F1} mm";
                    lblRefRot.Text = $"Góc xoay: {refe.RotDeg:F2}°";
                }
                else
                {
                    lblRefX.Text = "X: —";
                    lblRefY.Text = "Y: —";
                    lblRefZ.Text = "Z: —";
                    lblRefRot.Text = "Góc xoay: —";
                }
            }

            private void AddLogoIfFound(PictureBox logoBox)
            {
                try
                {
                    string logoPath = LogoHelper.GetLogoPath();
                    if (!string.IsNullOrEmpty(logoPath) && File.Exists(logoPath))
                        logoBox.Image = Image.FromFile(logoPath);
                }
                catch
                {
                }
            }

            private static double ParseDouble(string s)
            {
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out var v)) return v;
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;
                return 0.0;
            }
        }
    }
}
