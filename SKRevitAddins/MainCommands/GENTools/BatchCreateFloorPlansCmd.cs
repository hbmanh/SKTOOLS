using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ComboBox = System.Windows.Forms.ComboBox;
using Control = System.Windows.Forms.Control;
using Form = System.Windows.Forms.Form;
using Level = Autodesk.Revit.DB.Level;
using TextBox = System.Windows.Forms.TextBox;
using View = Autodesk.Revit.DB.View;

namespace SKRevitAddins.GENTools
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class BatchCreateFloorPlansCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string msg, ElementSet set)
        {
            UIApplication uiApp = c.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(lv => lv.Elevation)
                .ToList();

            if (levels.Count == 0)
            {
                TaskDialog.Show("Auto Create Plans", "Không tìm thấy Level nào trong dự án.");
                return Result.Failed;
            }

            var viewTemplates = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.IsTemplate &&
                           (v.ViewType == ViewType.FloorPlan ||
                            v.ViewType == ViewType.CeilingPlan ||
                            v.ViewType == ViewType.EngineeringPlan ||
                            v.ViewType == ViewType.AreaPlan))
                .OrderBy(v => v.Name)
                .ToList();

            using (var form = new LevelSelectForm(levels, viewTemplates))
            {
                if (form.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;

                string prefix = form.Prefix;
                string ending = form.Ending;
                List<Level> selected = form.SelectedLevels;
                View selectedTemplate = form.SelectedTemplate;
                string viewTypeStr = form.SelectedViewType;

                if (selected == null || selected.Count == 0)
                {
                    TaskDialog.Show("Auto Create Plans", "Chưa chọn Level nào để tạo mặt bằng.");
                    return Result.Cancelled;
                }

                // Xác định ViewFamily tương ứng
                ViewFamily selectedFamily = ViewFamily.FloorPlan;
                switch (viewTypeStr)
                {
                    case "Reflected Ceiling Plan":
                        selectedFamily = ViewFamily.CeilingPlan;
                        break;
                    case "Structural Plan":
                        selectedFamily = ViewFamily.StructuralPlan;
                        break;
                    case "Area Plan":
                        selectedFamily = ViewFamily.AreaPlan;
                        break;
                    default:
                        selectedFamily = ViewFamily.FloorPlan;
                        break;
                }

                ViewFamilyType viewType = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(vft => vft.ViewFamily == selectedFamily);

                if (viewType == null)
                {
                    TaskDialog.Show("Auto Create Plans", $"Không tìm thấy ViewFamilyType cho {viewTypeStr}.");
                    return Result.Failed;
                }

                using (Transaction t = new Transaction(doc, $"Auto Create {viewTypeStr}s"))
                {
                    t.Start();

                    int createdCount = 0;

                    foreach (var lv in selected)
                    {
                        string name = BuildViewName(prefix, lv.Name, ending);
                        if (ViewNameExists(doc, name)) continue;

                        ViewPlan vp = ViewPlan.Create(doc, viewType.Id, lv.Id);
                        vp.Name = name;

                        if (selectedTemplate != null)
                            vp.ViewTemplateId = selectedTemplate.Id;

                        createdCount++;
                    }

                    t.Commit();

                    TaskDialog.Show("Auto Create Plans",
                        $"Đã tạo {createdCount} {viewTypeStr}(s) mới.");
                }

                return Result.Succeeded;
            }
        }

        private string BuildViewName(string prefix, string level, string ending)
        {
            List<string> parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(prefix)) parts.Add(prefix.Trim());
            if (!string.IsNullOrWhiteSpace(level)) parts.Add(level.Trim());
            if (!string.IsNullOrWhiteSpace(ending)) parts.Add(ending.Trim());
            return string.Join(" - ", parts);
        }

        private bool ViewNameExists(Document doc, string name)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Any(v => v.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }
    }

    public class LevelSelectForm : Form
    {
        private CheckedListBox chkLevels;
        private TextBox txtPrefix;
        private TextBox txtEnding;
        private ComboBox cbTemplate;
        private ComboBox cbViewType; // combobox mới
        private Button btnOK, btnCancel;
        private PictureBox logo;
        private Label lblTitle;
        private List<Level> _levels;
        private List<View> _templates;

        public string Prefix => txtPrefix.Text;
        public string Ending => txtEnding.Text;
        public string SelectedViewType => cbViewType.SelectedItem?.ToString();
        public List<Level> SelectedLevels { get; private set; }
        public View SelectedTemplate { get; private set; }

        public LevelSelectForm(List<Level> levels, List<View> templates)
        {
            _levels = levels;
            _templates = templates;

            Text = "Auto Create Plans";
            Width = 440;
            Height = 630;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = System.Drawing.Color.White;

            // --- LOGO + TITLE ---
            logo = new PictureBox()
            {
                Left = 20,
                Top = 20,
                Width = 48,
                Height = 48,
                SizeMode = PictureBoxSizeMode.Zoom
            };
            try
            {
                logo.Image = System.Drawing.Image.FromFile(@"C:\ProgramData\SKRevitAddins\logo.png");
            }
            catch { }

            lblTitle = new Label()
            {
                Text = "Shinken Group®",
                Left = 80,
                Top = 32,
                Width = 300,
                Height = 30,
                Font = new System.Drawing.Font("Segoe UI", 13, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(25, 25, 25)
            };

            // --- INPUT FIELDS ---
            Label lbl1 = new Label() { Text = "Prefix:", Left = 20, Top = 85, Width = 70 };
            txtPrefix = new TextBox() { Left = 100, Top = 82, Width = 300 };

            Label lbl2 = new Label() { Text = "Ending:", Left = 20, Top = 115, Width = 70 };
            txtEnding = new TextBox() { Left = 100, Top = 112, Width = 300 };

            Label lbl3 = new Label() { Text = "View Template:", Left = 20, Top = 145, Width = 100 };
            cbTemplate = new ComboBox()
            {
                Left = 130,
                Top = 142,
                Width = 270,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            cbTemplate.Items.Add("<Không áp dụng>");
            foreach (var vt in templates)
                cbTemplate.Items.Add(vt.Name);
            cbTemplate.SelectedIndex = 0;

            // --- VIEW TYPE COMBOBOX ---
            Label lbl4 = new Label() { Text = "Loại mặt bằng:", Left = 20, Top = 175, Width = 100 };
            cbViewType = new ComboBox()
            {
                Left = 130,
                Top = 172,
                Width = 270,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cbViewType.Items.AddRange(new string[]
            {
                "Floor Plan",
                "Reflected Ceiling Plan",
                "Structural Plan",
                "Area Plan"
            });
            cbViewType.SelectedIndex = 0;

            // --- LEVEL LIST ---
            chkLevels = new CheckedListBox()
            {
                Left = 20,
                Top = 210,
                Width = 385,
                Height = 300,
                CheckOnClick = true
            };
            foreach (var lv in levels)
                chkLevels.Items.Add(lv.Name, true);

            // --- BUTTONS ---
            btnOK = new Button()
            {
                Text = "Tạo",
                Left = 230,
                Width = 80,
                Top = 530,
                DialogResult = DialogResult.OK
            };
            btnCancel = new Button()
            {
                Text = "Hủy",
                Left = 325,
                Width = 80,
                Top = 530,
                DialogResult = DialogResult.Cancel
            };

            Controls.AddRange(new Control[]
            {
                logo, lblTitle,
                lbl1, txtPrefix,
                lbl2, txtEnding,
                lbl3, cbTemplate,
                lbl4, cbViewType,
                chkLevels,
                btnOK, btnCancel
            });

            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                SelectedLevels = new List<Level>();
                foreach (var item in chkLevels.CheckedItems)
                {
                    var lv = _levels.FirstOrDefault(l => l.Name.Equals(item.ToString(), StringComparison.OrdinalIgnoreCase));
                    if (lv != null) SelectedLevels.Add(lv);
                }

                if (cbTemplate.SelectedIndex > 0)
                {
                    var vtName = cbTemplate.SelectedItem.ToString();
                    SelectedTemplate = _templates.FirstOrDefault(v => v.Name.Equals(vtName, StringComparison.OrdinalIgnoreCase));
                }
                else
                {
                    SelectedTemplate = null;
                }
            }
            base.OnFormClosing(e);
        }
    }

}
