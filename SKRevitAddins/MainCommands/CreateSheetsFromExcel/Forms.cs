using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Form = System.Windows.Forms.Form;
using Point = System.Drawing.Point;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;

namespace SKRevitAddins.CreateSheetsFromExcel
{
    public class ExcelSelectionForm : Form
    {
        public string SelectedFilePath { get; private set; }
        public bool CreateTemplate { get; private set; } = false;
        public string SelectedTitleBlock { get; private set; }

        public bool CreateWorkingView { get; private set; } = false;
        public string WorkingViewSuffix { get; private set; } = "";
        public bool CreateSheets { get; private set; } = true;
        public bool CreateSheetViews { get; private set; } = true;

        private ComboBox titleBlockCombo;
        private CheckBox cbSheetOnly, cbViewOnly, cbWorkingView;
        private TextBox txtSuffix;

        public ExcelSelectionForm(Document doc)
        {
            Text = "Shinken Group® - Tạo Sheet từ Excel";
            Size = new Size(700, 300);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            string logoPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "SKTools.bundle", "Icon", "shinken.png");

            PictureBox logo = new PictureBox
            {
                Image = File.Exists(logoPath) ? Image.FromFile(logoPath) : null,
                Size = new Size(48, 48),
                Location = new Point(20, 20),
                SizeMode = PictureBoxSizeMode.StretchImage
            };

            Label infoLabel = new Label
            {
                Text = "SHINKEN GROUP®\nChọn file Excel hoặc tạo mẫu mới để nhập thông tin Sheet.",
                Location = new Point(80, 30),
                AutoSize = true,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };

            Label titleBlockLabel = new Label
            {
                Text = "Chọn khung tên:",
                Location = new Point(40, 80),
                AutoSize = true
            };

            titleBlockCombo = new ComboBox
            {
                Location = new Point(160, 78),
                Size = new Size(500, 28),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            var titleBlocks = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilySymbol>()
                .Select(tb => $"{tb.Family.Name} : {tb.Name}")
                .OrderBy(n => n)
                .ToList();

            if (titleBlocks.Count == 0)
                titleBlocks.Add("Không có khung tên nào");

            titleBlockCombo.Items.AddRange(titleBlocks.ToArray());
            titleBlockCombo.SelectedIndex = 0;

            // ==== Options ====
            cbSheetOnly = new CheckBox { Text = "Tạo Sheet", Width = 200,Location = new Point(40, 120), Checked = true };
            cbViewOnly = new CheckBox { Text = "Tạo Sheet's View", Width = 200, Location = new Point(40, 150), Checked = true };
            cbWorkingView = new CheckBox { Text = "Tạo Working's View (hậu tố):", Width = 200,Location = new Point(40, 180) };
            txtSuffix = new TextBox { Location = new Point(250, 178), Width = 120, Enabled = false };

            cbWorkingView.CheckedChanged += (s, e) =>
            {
                txtSuffix.Enabled = cbWorkingView.Checked;
            };

            // ==== Buttons ====
            Button btnChoose = new Button
            {
                Text = "Chọn file Excel...",
                Size = new Size(180, 30),
                Location = new Point(40, 220),
                TextAlign = ContentAlignment.MiddleCenter
            };
            btnChoose.Click += (s, e) =>
            {
                OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel files (*.xlsx)|*.xlsx" };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    SelectedFilePath = ofd.FileName;
                    SetFormValues();
                    DialogResult = DialogResult.OK;
                    Close();
                }
            };

            Button btnCreate = new Button
            {
                Text = "Tạo mẫu Excel",
                Size = new Size(180, 30),
                Location = new Point(240, 220),
                TextAlign = ContentAlignment.MiddleCenter
            };
            btnCreate.Click += (s, e) =>
            {
                CreateTemplate = true;
                SetFormValues();
                DialogResult = DialogResult.Yes;
                Close();
            };

            Controls.Add(logo);
            Controls.Add(infoLabel);
            Controls.Add(titleBlockLabel);
            Controls.Add(titleBlockCombo);
            Controls.Add(cbSheetOnly);
            Controls.Add(cbViewOnly);
            Controls.Add(cbWorkingView);
            Controls.Add(txtSuffix);
            Controls.Add(btnChoose);
            Controls.Add(btnCreate);
        }

        private void SetFormValues()
        {
            SelectedTitleBlock = titleBlockCombo.SelectedItem?.ToString();
            CreateWorkingView = cbWorkingView.Checked;
            WorkingViewSuffix = txtSuffix.Text.Trim();
            CreateSheets = cbSheetOnly.Checked;
            CreateSheetViews = cbViewOnly.Checked;
        }
    }

    public class ProgressForm : Form
    {
        public ProgressBar progressBar;
        private Label sheetLabel;
        private Button cancelButton;
        public bool IsCanceled { get; private set; } = false;

        public ProgressForm(int maxValue)
        {
            Text = "Shinken Group® - Đang xử lý";
            Size = new Size(480, 160);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            PictureBox logo = new PictureBox
            {
                Image = Image.FromFile("C:\\ProgramData\\Autodesk\\Revit\\Addins\\2023\\SKTools.bundle\\Contents\\Resources\\Images\\shinken.png"),
                Size = new Size(32, 32),
                Location = new Point(10, 10),
                SizeMode = PictureBoxSizeMode.StretchImage
            };

            sheetLabel = new Label
            {
                Text = "Đang xử lý sheet...",
                Location = new Point(50, 15),
                AutoSize = true,
                Font = new Font("Segoe UI", 9)
            };

            progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Continuous,
                Maximum = maxValue,
                Minimum = 0,
                Value = 0,
                Location = new Point(50, 50),
                Size = new Size(380, 25)
            };

            cancelButton = new Button
            {
                Text = "Hủy",
                Size = new Size(80, 30),
                Location = new Point(350, 90)
            };
            cancelButton.Click += (s, e) =>
            {
                IsCanceled = true;
                cancelButton.Enabled = false;
                sheetLabel.Text = "Đang dừng...";
            };

            Controls.Add(logo);
            Controls.Add(sheetLabel);
            Controls.Add(progressBar);
            Controls.Add(cancelButton);
        }

        public void UpdateProgress(int value, string currentSheetName)
        {
            progressBar.Value = Math.Min(value, progressBar.Maximum);
            int percent = (int)((value / (float)progressBar.Maximum) * 100);
            sheetLabel.Text = $"[{percent}%] Đang xử lý: {currentSheetName}";
            Refresh();
        }
    }
}
