using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Form = System.Windows.Forms.Form;
using Point = System.Drawing.Point;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;

namespace SKRevitAddins.Commands.CreateSheetsFromExcel
{
    public class ExcelSelectionForm : Form
    {
        public string SelectedFilePath { get; private set; }
        public bool CreateTemplate { get; private set; } = false;
        public string SelectedTitleBlock { get; private set; }

        private ComboBox titleBlockCombo;

        public ExcelSelectionForm(Document doc)
        {
            Text = "Shinken Group® - Tạo Sheet từ Excel";
            Size = new Size(480, 240);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            PictureBox logo = new PictureBox
            {
                Image = Image.FromFile("C:\\ProgramData\\Autodesk\\Revit\\Addins\\2023\\SKTools.bundle\\Contents\\Resources\\Images\\shinken.png"),
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
                Location = new Point(160, 80),
                Size = new Size(260, 28),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // Lấy danh sách FamilySymbol thuộc TitleBlock
            var titleBlocks = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilySymbol>()
                .Select(tb => $"{tb.Family.Name} : {tb.Name}")
                .OrderBy(n => n)
                .ToList();

            if (titleBlocks.Count == 0)
            {
                titleBlocks.Add("Không có khung tên nào");
            }

            titleBlockCombo.Items.AddRange(titleBlocks.ToArray());
            titleBlockCombo.SelectedIndex = 0;

            Button btnChoose = new Button
            {
                Text = "Chọn file\nExcel...",
                Size = new Size(180, 40),
                Location = new Point(40, 130),
                TextAlign = ContentAlignment.MiddleCenter
            };
            btnChoose.Click += (s, e) =>
            {
                OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel files (*.xlsx)|*.xlsx" };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    SelectedFilePath = ofd.FileName;
                    SelectedTitleBlock = titleBlockCombo.SelectedItem?.ToString();
                    DialogResult = DialogResult.OK;
                    Close();
                }
            };

            Button btnCreate = new Button
            {
                Text = "Tạo mẫu\nExcel",
                Size = new Size(180, 40),
                Location = new Point(240, 130),
                TextAlign = ContentAlignment.MiddleCenter
            };
            btnCreate.Click += (s, e) =>
            {
                CreateTemplate = true;
                SelectedTitleBlock = titleBlockCombo.SelectedItem?.ToString();
                DialogResult = DialogResult.Yes;
                Close();
            };

            Controls.Add(logo);
            Controls.Add(infoLabel);
            Controls.Add(titleBlockLabel);
            Controls.Add(titleBlockCombo);
            Controls.Add(btnChoose);
            Controls.Add(btnCreate);
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
