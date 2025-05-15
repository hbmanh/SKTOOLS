// EquipmentClassificationProcessor.cs
// Revit API (.NET Framework 4.8)
// 2025‑05‑04  – Loại bỏ instance có Equipment ID‑SP = BS hoặc STR (không xóa model)
//             Đã sửa lỗi CS1503: ép kiểu Element → FamilyInstance trong UpdateCount()

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Form = System.Windows.Forms.Form;

namespace SKRevitAddins
{
    [Transaction(TransactionMode.Manual)]
    public class EquipmentClassificationProcessor : IExternalCommand
    {
        #region ---------- helpers ----------
        private static void ClearTextParam(Parameter p)
        {
            if (p != null && !p.IsReadOnly && p.StorageType == StorageType.String)
                p.Set(string.Empty);
        }

        private static void WriteText(Parameter p, string value)
        {
            if (p != null && !p.IsReadOnly && p.StorageType == StorageType.String)
                p.Set(value);
        }

        private static bool IsCU(string s)
        {
            // khớp "CU" nguyên từ, không khớp "FCU", "ACU" …
            return Regex.IsMatch(s ?? string.Empty,
                                 @"(^|\W)CU($|\W)",
                                 RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// True nếu FamilyInstance có Equipment ID‑SP = BS hoặc STR.
        /// </summary>
        private static bool IsExcludedById(FamilyInstance fi)
        {
            string eid = fi.LookupParameter("Equipment ID-SP")?.AsString()?.Trim();
            return string.Equals(eid, "BS", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(eid, "STR", StringComparison.OrdinalIgnoreCase);
        }
        #endregion

        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // --------- 1. chọn Category ----------
            var defaultCats = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_Furniture,
                BuiltInCategory.OST_SpecialityEquipment,
                BuiltInCategory.OST_DuctTerminal
            };

            using var form = new CategoryMultiSelectForm(defaultCats, doc);
            if (form.ShowDialog() != DialogResult.OK ||
                form.SelectedCategories.Count == 0)
                return Result.Cancelled;

            // --------- 2. thu thập instance ----------
            IList<FamilyInstance> instances = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => form.SelectedCategories
                               .Contains((BuiltInCategory)fi.Category.Id.IntegerValue))
                .Where(fi => !IsExcludedById(fi))          // bỏ BS/STR
                .ToList();

            // --------- 3. gom nhóm 7 yếu tố ----------
            var groups = new Dictionary<
                (string cls, string eid, string level, string schedGrp,
                 string model, ElementId typeId, string tagNote),
                List<FamilyInstance>>();

            foreach (FamilyInstance fi in instances)
            {
                FamilySymbol sym = fi.Symbol;

                // Classification (type parameter) – bắt buộc
                Parameter clsP = sym.LookupParameter("Equipment Classification-SP");
                if (clsP == null || clsP.StorageType != StorageType.String) continue;

                string cls = clsP.AsString()?.Trim();
                if (string.IsNullOrWhiteSpace(cls)) continue;

                // Các tham số còn lại – cho phép rỗng
                string eid = fi.LookupParameter("Equipment ID-SP")?.AsString()?.Trim() ?? string.Empty;
                string level = fi.LookupParameter("Level-SP")?.AsString()?.Trim() ?? string.Empty;
                string sched = fi.LookupParameter("SChedule Group-SP")?.AsString()?.Trim() ?? string.Empty;
                string tagNote = fi.LookupParameter("Unit Tag Note-SP")?.AsString()?.Trim() ?? string.Empty;

                // ---- Model (type parameter) ----
                Parameter modelP = sym.LookupParameter("Model") ??
                                   sym.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL);
                string model = modelP?.AsString()?.Trim() ?? string.Empty;

                var key = (cls, eid, level, sched, model, sym.Id, tagNote);

                if (!groups.ContainsKey(key))
                    groups[key] = new List<FamilyInstance>();

                groups[key].Add(fi);
            }

            // --------- 4. transaction ----------
            using (var tx = new Transaction(doc, "Update Indoor/Outdoor Qty"))
            {
                tx.Start();

                // B1: reset
                foreach (FamilyInstance fi in instances)
                {
                    ClearTextParam(fi.LookupParameter("Q'ty Indoor Unit-SP"));
                    ClearTextParam(fi.LookupParameter("Q'ty Outdoor Unit-SP"));
                }

                // B2: đếm & ghi
                foreach (var kvp in groups)
                {
                    string cls = kvp.Key.cls;
                    int qty = kvp.Value.Count;
                    bool isCu = IsCU(cls);

                    foreach (FamilyInstance fi in kvp.Value)
                    {
                        Parameter indoor = fi.LookupParameter("Q'ty Indoor Unit-SP");
                        Parameter outdoor = fi.LookupParameter("Q'ty Outdoor Unit-SP");

                        if (isCu)          // Condensing Unit
                        {
                            WriteText(indoor, "-");
                            WriteText(outdoor, qty.ToString());
                        }
                        else               // thiết bị khác
                        {
                            WriteText(indoor, qty.ToString());
                            WriteText(outdoor, "-");
                        }
                    }
                }

                tx.Commit();
            }

            TaskDialog.Show("Hoàn tất",
                $"Đã reset & cập nhật {instances.Count} thiết bị (gộp theo 7 yếu tố).");
            return Result.Succeeded;
        }

        // ---------- UI chọn nhiều Category ----------
        private class CategoryMultiSelectForm : Form
        {
            private readonly CheckedListBox _clb;
            private readonly Label _lbl;
            private readonly Button _ok;
            private readonly Document _doc;

            public List<BuiltInCategory> SelectedCategories { get; } = new();

            public CategoryMultiSelectForm(IEnumerable<BuiltInCategory> cats, Document doc)
            {
                _doc = doc;
                Text = "Chọn Category cần xử lý";
                Width = 340;
                Height = 380;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                StartPosition = FormStartPosition.CenterScreen;

                _clb = new CheckedListBox { Dock = DockStyle.Top, Height = 240 };
                _lbl = new Label { Dock = DockStyle.Top, Height = 26, Text = "Số lượng: 0" };
                _ok = new Button { Dock = DockStyle.Bottom, Text = "Thực hiện" };

                foreach (BuiltInCategory bic in cats)
                {
                    string name = _doc.Settings.Categories.get_Item(bic).Name;
                    _clb.Items.Add(new CatItem(bic, name), false);
                }

                _clb.ItemCheck += (_, __) => BeginInvoke((MethodInvoker)UpdateCount);
                _ok.Click += (_, __) =>
                {
                    foreach (CatItem ci in _clb.CheckedItems)
                        SelectedCategories.Add(ci.BIC);
                    DialogResult = DialogResult.OK;
                    Close();
                };

                Controls.Add(_clb);
                Controls.Add(_lbl);
                Controls.Add(_ok);
            }

            private void UpdateCount()
            {
                var selected = _clb.CheckedItems.Cast<CatItem>()
                                 .Select(ci => ci.BIC)
                                 .ToHashSet();

                int n = new FilteredElementCollector(_doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()                        // ép kiểu để dùng IsExcludedById
                    .Count(fi => selected
                                 .Contains((BuiltInCategory)fi.Category.Id.IntegerValue)
                              && !IsExcludedById(fi));              // bỏ BS/STR

                _lbl.Text = $"Số lượng: {n}";
            }

            private class CatItem
            {
                public BuiltInCategory BIC { get; }
                private readonly string _name;
                public CatItem(BuiltInCategory bic, string name)
                {
                    BIC = bic;
                    _name = name;
                }
                public override string ToString() => _name;
            }
        }
    }
}
