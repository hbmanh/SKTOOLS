using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using SKRevitAddins.Utils;
using ComboBox = System.Windows.Forms.ComboBox;
using Control = System.Windows.Forms.Control;
using Form = System.Windows.Forms.Form;
using Level = Autodesk.Revit.DB.Level;
using TextBox = System.Windows.Forms.TextBox;
using View = Autodesk.Revit.DB.View;
using Color = System.Drawing.Color;
using Panel = System.Windows.Forms.Panel;

namespace SKRevitAddins.GENTools
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class BatchCreateOrDeletePlansCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string msg, ElementSet set)
        {
            UIApplication uiApp = c.Application;
            if (uiApp == null)
            {
                TaskDialog.Show("Auto Create/Delete Plans", "Không có UIApplication.");
                return Result.Failed;
            }

            UIDocument uiDoc = uiApp.ActiveUIDocument;
            if (uiDoc == null)
            {
                TaskDialog.Show("Auto Create/Delete Plans", "Không có tài liệu hiện hành (Active Document).");
                return Result.Failed;
            }

            Document doc = uiDoc.Document;

            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(lv => lv.Elevation)
                .ToList();

            var viewTemplates = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.IsTemplate)
                .OrderBy(v => v.Name)
                .ToList();

            using (var main = new MainForm(levels, viewTemplates, doc))
            {
                main.ShowDialog();
            }

            return Result.Succeeded;
        }

        private static string BuildViewName(string prefix, string level, string ending)
        {
            List<string> parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(prefix)) parts.Add(prefix.Trim());
            if (!string.IsNullOrWhiteSpace(level)) parts.Add(level.Trim());
            if (!string.IsNullOrWhiteSpace(ending)) parts.Add(ending.Trim());
            return string.Join(" - ", parts);
        }

        private static bool ViewNameExists(Document doc, string name)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Any(v => v.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        public class MainForm : Form
        {
            private Button btnCreate;
            private Button btnDelete;
            private PictureBox logo;
            private Label lblTitle;
            private readonly List<Level> _levels;
            private readonly List<View> _templates;
            private readonly Document _doc;

            public MainForm(List<Level> levels, List<View> templates, Document doc)
            {
                _levels = levels;
                _templates = templates;
                _doc = doc;

                Text = "SKRevit Tools";
                MinimumSize = new Size(360, 220);
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                AutoScaleMode = AutoScaleMode.Dpi;
                BackColor = Color.White;
                Font = SystemFonts.MessageBoxFont;

                var mainTable = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = 3,
                    Padding = new Padding(12),
                    BackColor = Color.White
                };
                mainTable.RowStyles.Clear();
                mainTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                mainTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var headerFlow = new FlowLayoutPanel
                {
                    Dock = DockStyle.Top,
                    FlowDirection = FlowDirection.LeftToRight,
                    AutoSize = true,
                    WrapContents = false,
                    Margin = new Padding(0, 0, 0, 12)
                };

                logo = new PictureBox
                {
                    Width = 48,
                    Height = 48,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Margin = new Padding(0, 0, 12, 0)
                };
                AddLogoIfFound(logo);

                lblTitle = new Label
                {
                    Text = "SKRevit Addins",
                    AutoSize = true,
                    Font = new Font("Segoe UI", 13, FontStyle.Bold),
                    ForeColor = Color.FromArgb(25, 25, 25),
                    TextAlign = ContentAlignment.MiddleCenter
                };

                headerFlow.Controls.Add(logo);
                headerFlow.Controls.Add(lblTitle);

                var btnsFlow = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.TopDown,
                    AutoSize = true
                };

                btnCreate = new Button
                {
                    Text = "Auto Create Plans",
                    AutoSize = true,
                    Padding = new Padding(8),
                    Margin = new Padding(0, -20, 0, 8)
                };
                btnCreate.Click += BtnCreate_Click;

                btnDelete = new Button
                {
                    Text = "Delete Plans / Views",
                    AutoSize = true,
                    Padding = new Padding(8),
                    Margin = new Padding(0, -5, 0, 16)
                };
                btnDelete.Click += BtnDelete_Click;

                btnsFlow.Controls.Add(btnCreate);
                btnsFlow.Controls.Add(btnDelete);

                mainTable.Controls.Add(headerFlow, 0, 0);
                mainTable.Controls.Add(new Panel { Height = 8, Dock = DockStyle.Top }, 0, 1);
                mainTable.Controls.Add(btnsFlow, 0, 2);

                Controls.Add(mainTable);
                AcceptButton = btnCreate;
            }

            private void AddLogoIfFound(PictureBox logoBox)
            {
                try
                {
                    string logoPath = LogoHelper.GetLogoPath();
                    if (!string.IsNullOrEmpty(logoPath) && File.Exists(logoPath))
                    {
                        logoBox.Image = Image.FromFile(logoPath);
                    }
                }
                catch
                {
                }
            }

            private void BtnCreate_Click(object sender, EventArgs e)
            {
                using (var form = new LevelSelectForm(_levels, _templates))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                        return;

                    string prefix = form.Prefix;
                    string ending = form.Ending;
                    List<Level> selected = form.SelectedLevels;
                    View selectedTemplate = form.SelectedTemplate;
                    ViewTypeItem selectedViewTypeItem = form.SelectedViewTypeItem;

                    if (selected == null || selected.Count == 0)
                    {
                        TaskDialog.Show("Auto Create Plans", "Chưa chọn Level nào để tạo mặt bằng.");
                        return;
                    }

                    ViewFamily selectedFamily = selectedViewTypeItem.Family;

                    ViewFamilyType viewType = new FilteredElementCollector(_doc)
                        .OfClass(typeof(ViewFamilyType))
                        .Cast<ViewFamilyType>()
                        .FirstOrDefault(vft => vft.ViewFamily == selectedFamily);

                    if (viewType == null)
                    {
                        TaskDialog.Show("Auto Create Plans", $"Không tìm thấy ViewFamilyType cho {selectedViewTypeItem.Name}.");
                        return;
                    }

                    using (Transaction t = new Transaction(_doc, $"Auto Create {selectedViewTypeItem.Name}s"))
                    {
                        try
                        {
                            t.Start();
                            int createdCount = 0;
                            foreach (var lv in selected)
                            {
                                string name = BuildViewName(prefix, lv.Name, ending);
                                if (ViewNameExists(_doc, name)) continue;

                                ViewPlan vp = ViewPlan.Create(_doc, viewType.Id, lv.Id);
                                if (vp == null) continue;
                                vp.Name = name;

                                if (selectedTemplate != null)
                                {
                                    if (selectedTemplate.ViewType == vp.ViewType)
                                        vp.ViewTemplateId = selectedTemplate.Id;
                                }

                                createdCount++;
                            }
                            t.Commit();
                            TaskDialog.Show("Auto Create Plans", $"Đã tạo {createdCount} {selectedViewTypeItem.Name}(s) mới.");
                        }
                        catch (Exception ex)
                        {
                            if (t.GetStatus() == TransactionStatus.Started)
                                t.RollBack();
                            TaskDialog.Show("Auto Create Plans - Lỗi", $"Đã xảy ra lỗi: {ex.Message}");
                        }
                    }
                }
            }

            private void BtnDelete_Click(object sender, EventArgs e)
            {
                using (var form = new DeleteViewsForm(_doc))
                {
                    form.ShowDialog();
                }
            }
        }

        public class ViewTypeItem
        {
            public string Name { get; set; }
            public ViewFamily Family { get; set; }
            public override string ToString() => Name;
        }

        public class LevelSelectForm : Form
        {
            private CheckedListBox chkLevels;
            private TextBox txtPrefix;
            private TextBox txtEnding;
            private ComboBox cbTemplate;
            private ComboBox cbViewType;
            private Button btnOK;
            private Button btnCancel;
            private PictureBox logo;
            private Label lblTitle;
            private readonly List<Level> _levels;
            private readonly List<View> _templates;

            public string Prefix => txtPrefix.Text;
            public string Ending => txtEnding.Text;
            public ViewTypeItem SelectedViewTypeItem => cbViewType.SelectedItem as ViewTypeItem;
            public List<Level> SelectedLevels { get; private set; }
            public View SelectedTemplate { get; private set; }

            public LevelSelectForm(List<Level> levels, List<View> templates)
            {
                _levels = levels;
                _templates = templates;

                Text = "Auto Create Plans";
                MinimumSize = new Size(840, 620);
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                AutoScaleMode = AutoScaleMode.Dpi;
                BackColor = Color.White;
                Font = SystemFonts.MessageBoxFont;

                var mainTable = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = 4,
                    Padding = new Padding(12)
                };
                mainTable.RowStyles.Clear();
                mainTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                mainTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var headerFlow = new FlowLayoutPanel
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    FlowDirection = FlowDirection.LeftToRight
                };
                logo = new PictureBox
                {
                    Width = 48,
                    Height = 48,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Margin = new Padding(0, 0, 12, 0)
                };
                AddLogoIfFound(logo);
                lblTitle = new Label
                {
                    Text = "Shinken Group®",
                    AutoSize = true,
                    Font = new Font("Segoe UI", 13, FontStyle.Bold),
                    TextAlign = ContentAlignment.MiddleLeft
                };
                headerFlow.Controls.Add(logo);
                headerFlow.Controls.Add(lblTitle);

                var inputsTable = new TableLayoutPanel
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    ColumnCount = 4
                };
                inputsTable.ColumnStyles.Clear();
                inputsTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                inputsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
                inputsTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                inputsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

                inputsTable.Controls.Add(new Label { Text = "Prefix:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
                txtPrefix = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
                inputsTable.Controls.Add(txtPrefix, 1, 0);

                inputsTable.Controls.Add(new Label { Text = "Ending:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 0);
                txtEnding = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
                inputsTable.Controls.Add(txtEnding, 3, 0);

                inputsTable.Controls.Add(new Label { Text = "View Template:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
                cbTemplate = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left | AnchorStyles.Right };
                cbTemplate.Items.Add("<Không áp dụng>");
                foreach (var vt in _templates) cbTemplate.Items.Add(vt.Name);
                cbTemplate.SelectedIndex = 0;
                inputsTable.Controls.Add(cbTemplate, 1, 1);

                inputsTable.Controls.Add(new Label { Text = "Loại mặt bằng:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 1);
                cbViewType = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left | AnchorStyles.Right };
                cbViewType.Items.Add(new ViewTypeItem { Name = "Floor Plan", Family = ViewFamily.FloorPlan });
                cbViewType.Items.Add(new ViewTypeItem { Name = "Reflected Ceiling Plan", Family = ViewFamily.CeilingPlan });
                cbViewType.Items.Add(new ViewTypeItem { Name = "Structural Plan", Family = ViewFamily.StructuralPlan });
                cbViewType.Items.Add(new ViewTypeItem { Name = "Area Plan", Family = ViewFamily.AreaPlan });
                cbViewType.SelectedIndex = 0;
                inputsTable.Controls.Add(cbViewType, 3, 1);

                chkLevels = new CheckedListBox
                {
                    Dock = DockStyle.Fill,
                    CheckOnClick = true
                };
                foreach (var lv in _levels) chkLevels.Items.Add(lv.Name, true);

                var buttonsFlow = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.RightToLeft,
                    AutoSize = true
                };
                btnOK = new Button { Text = "Tạo", AutoSize = true, DialogResult = DialogResult.OK };
                btnCancel = new Button { Text = "Hủy", AutoSize = true, DialogResult = DialogResult.Cancel };
                buttonsFlow.Controls.Add(btnOK);
                buttonsFlow.Controls.Add(btnCancel);

                mainTable.Controls.Add(headerFlow, 0, 0);
                mainTable.Controls.Add(inputsTable, 0, 1);
                mainTable.Controls.Add(chkLevels, 0, 2);
                mainTable.Controls.Add(buttonsFlow, 0, 3);

                Controls.Add(mainTable);
                AcceptButton = btnOK;
                CancelButton = btnCancel;
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

        public class DeleteViewsForm : Form
        {
            private readonly Document _doc;
            private CheckedListBox chkViews;
            private ComboBox cbFilterType;
            private TextBox txtSearch;
            private Button btnRefresh;
            private Button btnDelete;
            private Button btnCancel;
            private Label lblCount;
            private Button btnSelectAll;
            private Button btnUnselectAll;

            private List<View> allViews;

            public DeleteViewsForm(Document doc)
            {
                _doc = doc;

                Text = "Delete Views - SKRevit";
                MinimumSize = new Size(1000, 480);
                StartPosition = FormStartPosition.CenterScreen;
                AutoScaleMode = AutoScaleMode.Dpi;
                Font = SystemFonts.MessageBoxFont;

                var mainTable = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = 3,
                    Padding = new Padding(8),
                    AutoSize = false
                };
                mainTable.RowStyles.Clear();
                mainTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                mainTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var topFlow = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.LeftToRight,
                    WrapContents = false,
                    AutoSize = true,
                    Margin = new Padding(0, 0, 0, 6)
                };

                var lblFilter = new Label
                {
                    Text = "Filter (View Type):",
                    AutoSize = true,
                    TextAlign = ContentAlignment.MiddleLeft
                };
                cbFilterType = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Width = 200,
                    Anchor = AnchorStyles.Left | AnchorStyles.Top
                };
                cbFilterType.Items.Add("All");
                cbFilterType.Items.Add("FloorPlan");
                cbFilterType.Items.Add("CeilingPlan");
                cbFilterType.Items.Add("StructuralPlan");
                cbFilterType.Items.Add("AreaPlan");
                cbFilterType.Items.Add("DraftingView");
                cbFilterType.Items.Add("ThreeD");
                cbFilterType.Items.Add("Elevation");
                cbFilterType.Items.Add("Schedule");
                cbFilterType.SelectedIndex = 0;
                cbFilterType.SelectedIndexChanged += Filter_Changed;

                var lblSearch = new Label
                {
                    Text = "Search (contains):",
                    AutoSize = true,
                    Margin = new Padding(12, 6, 6, 0)
                };
                txtSearch = new TextBox
                {
                    Width = 220,
                    Anchor = AnchorStyles.Left | AnchorStyles.Top
                };
                txtSearch.TextChanged += Filter_Changed;

                btnRefresh = new Button
                {
                    Text = "Refresh",
                    AutoSize = true,
                    Anchor = AnchorStyles.Left | AnchorStyles.Top
                };
                btnRefresh.Click += BtnRefresh_Click;

                topFlow.Controls.Add(lblFilter);
                topFlow.Controls.Add(cbFilterType);
                topFlow.Controls.Add(lblSearch);
                topFlow.Controls.Add(txtSearch);
                topFlow.Controls.Add(btnRefresh);

                chkViews = new CheckedListBox
                {
                    Dock = DockStyle.Fill,
                    CheckOnClick = true,
                    HorizontalScrollbar = true
                };

                var listPanel = new Panel
                {
                    Dock = DockStyle.Fill,
                    Padding = new Padding(0)
                };
                chkViews.Parent = listPanel;
                chkViews.Dock = DockStyle.Fill;

                var bottomFlow = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.LeftToRight,
                    WrapContents = false,
                    AutoSize = true,
                    Margin = new Padding(0, 6, 0, 0)
                };

                btnSelectAll = new Button { Text = "Select all", AutoSize = true };
                btnSelectAll.Click += BtnSelectAll_Click;

                btnUnselectAll = new Button { Text = "Unselect all", AutoSize = true };
                btnUnselectAll.Click += BtnUnselectAll_Click;

                lblCount = new Label
                {
                    AutoSize = true,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Margin = new Padding(12, 8, 6, 6)
                };

                var rightActions = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.LeftToRight,
                    AutoSize = true,
                    WrapContents = false,
                    Anchor = AnchorStyles.Right,
                    Margin = new Padding(0)
                };

                btnDelete = new Button { Text = "Delete Selected", AutoSize = true };
                btnDelete.Click += BtnDelete_Click;
                btnCancel = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.Cancel };

                rightActions.Controls.Add(btnDelete);
                rightActions.Controls.Add(btnCancel);

                bottomFlow.Controls.Add(btnSelectAll);
                bottomFlow.Controls.Add(btnUnselectAll);
                bottomFlow.Controls.Add(lblCount);

                var filler = new Panel { AutoSize = false, Width = 20, Dock = DockStyle.Fill };
                bottomFlow.Controls.Add(filler);
                bottomFlow.Controls.Add(rightActions);

                mainTable.Controls.Add(topFlow, 0, 0);
                mainTable.Controls.Add(listPanel, 0, 1);
                mainTable.Controls.Add(bottomFlow, 0, 2);

                Controls.Add(mainTable);

                LoadAllViews();
                PopulateList();

                txtSearch.TabIndex = 0;
                chkViews.TabIndex = 1;
            }

            private void LoadAllViews()
            {
                allViews = new FilteredElementCollector(_doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate)
                    .OrderBy(v => v.ViewType.ToString())
                    .ThenBy(v => v.Name)
                    .ToList();
            }

            private void PopulateList()
            {
                chkViews.Items.Clear();
                string filterType = cbFilterType.SelectedItem?.ToString() ?? "All";
                string search = txtSearch.Text ?? string.Empty;
                var filtered = allViews.AsEnumerable();

                if (!string.IsNullOrWhiteSpace(filterType) && filterType != "All")
                    filtered = filtered.Where(v => v.ViewType.ToString().Equals(filterType, StringComparison.OrdinalIgnoreCase));

                if (!string.IsNullOrWhiteSpace(search))
                    filtered = filtered.Where(v => v.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0);

                foreach (var v in filtered)
                {
                    string display = $"{v.Name}  [{v.ViewType}]";
                    chkViews.Items.Add(display, false);
                }

                lblCount.Text = $"Found: {chkViews.Items.Count} views (Templates ignored).";
            }

            private void Filter_Changed(object sender, EventArgs e)
            {
                PopulateList();
            }

            private void BtnRefresh_Click(object sender, EventArgs e)
            {
                LoadAllViews();
                PopulateList();
            }

            private void BtnSelectAll_Click(object sender, EventArgs e)
            {
                SetAllChecked(true);
            }

            private void BtnUnselectAll_Click(object sender, EventArgs e)
            {
                SetAllChecked(false);
            }

            private void SetAllChecked(bool check)
            {
                for (int i = 0; i < chkViews.Items.Count; i++)
                    chkViews.SetItemChecked(i, check);
            }

            private void BtnDelete_Click(object sender, EventArgs e)
            {
                if (chkViews.CheckedItems.Count == 0)
                {
                    MessageBox.Show("Chưa chọn view nào để xóa.", "Delete Views", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var confirm = TaskDialog.Show(
                    "Delete Views",
                    $"Bạn có chắc muốn xóa {chkViews.CheckedItems.Count} view đã chọn? (Hành động này không thể hoàn tác nếu commit)",
                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                if (confirm != TaskDialogResult.Yes) return;

                var namesToDelete = new List<string>();
                foreach (var it in chkViews.CheckedItems)
                {
                    string s = it.ToString();
                    int idx = s.LastIndexOf("  [", StringComparison.Ordinal);
                    string nameOnly = idx >= 0 ? s.Substring(0, idx) : s;
                    namesToDelete.Add(nameOnly.Trim());
                }

                var viewsToDelete = allViews.Where(v => namesToDelete.Contains(v.Name)).ToList();

                if (viewsToDelete.Count == 0)
                {
                    MessageBox.Show("Không tìm thấy view tương ứng để xóa.", "Delete Views", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int success = 0;
                int failed = 0;
                List<string> failedNames = new List<string>();

                using (Transaction t = new Transaction(_doc, "Delete Selected Views"))
                {
                    try
                    {
                        t.Start();
                        foreach (var v in viewsToDelete)
                        {
                            try
                            {
                                if (!v.IsValidObject)
                                {
                                    failed++;
                                    failedNames.Add(v.Name + " (invalid)");
                                    continue;
                                }
                                if (_doc.IsFamilyDocument)
                                {
                                    failed++;
                                    failedNames.Add(v.Name + " (family doc)");
                                    continue;
                                }

                                View active = _doc.ActiveView;
                                if (active != null && active.Id == v.Id)
                                {
                                    failed++;
                                    failedNames.Add(v.Name + " (active view)");
                                    continue;
                                }

                                _doc.Delete(v.Id);
                                success++;
                            }
                            catch (Exception exInner)
                            {
                                failed++;
                                failedNames.Add(v.Name + " (" + exInner.Message + ")");
                            }
                        }
                        t.Commit();
                    }
                    catch (Exception ex)
                    {
                        if (t.GetStatus() == TransactionStatus.Started)
                            t.RollBack();
                        MessageBox.Show("Lỗi khi xóa views: " + ex.Message, "Delete Views", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                string resultMsg = $"Hoàn tất. Deleted: {success}. Failed: {failed}.";
                if (failedNames.Count > 0)
                {
                    resultMsg += Environment.NewLine + "Failed list (một số lý do):" +
                                 Environment.NewLine + string.Join(Environment.NewLine, failedNames.Take(50));
                }

                MessageBox.Show(resultMsg, "Delete Views", MessageBoxButtons.OK, MessageBoxIcon.Information);

                LoadAllViews();
                PopulateList();
            }
        }
    }
}
