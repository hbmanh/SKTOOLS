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
                Text = "Chọn file" + Environment.NewLine + "Excel...",
                Size = new Size(180, 40),
                Location = new Point(40, 120),
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
                Text = "Tạo mẫu" + Environment.NewLine + "Excel",
                Size = new Size(180, 40),
                Location = new Point(240, 120),
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

        public ProgressForm(int maxValue)
        {
            Text = "Shinken Group® - Đang xử lý";
            Size = new Size(420, 100);
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

            progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Continuous,
                Maximum = maxValue,
                Minimum = 0,
                Value = 0,
                Location = new Point(50, 20),
                Size = new Size(340, 25)
            };

            Controls.Add(logo);
            Controls.Add(progressBar);
        }

        public void UpdateProgress(int value)
        {
            progressBar.Value = value;
            Refresh();
        }
    }

    public class SheetSelectionForm : Form
    {
        public List<string> SelectedSheets = new List<string>();
        private DataGridView sheetGrid;

        public SheetSelectionForm(List<(string number, string name)> existingSheets)
        {
            Text = "Shinken Group® - Chọn sheet để tạo lại";
            Size = new Size(500, 540);
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

            Label label = new Label
            {
                Text = "Các sheet đã tồn tại, chọn để tạo lại:",
                Location = new Point(50, 15),
                AutoSize = true
            };

            sheetGrid = new DataGridView
            {
                Location = new Point(20, 50),
                Size = new Size(440, 320),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                MultiSelect = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false
            };

            sheetGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Sheet Number", Width = 150 });
            sheetGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Sheet Name", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });

            foreach (var sheet in existingSheets)
                sheetGrid.Rows.Add(sheet.number, sheet.name);

            sheetGrid.ClearSelection();

            Button btnSelectAll = new Button
            {
                Text = "Chọn tất cả",
                Width = 100,
                Height = 30
            };
            btnSelectAll.Click += (s, e) => sheetGrid.SelectAll();

            Button btnDeselectAll = new Button
            {
                Text = "Bỏ chọn",
                Width = 100,
                Height = 30
            };
            btnDeselectAll.Click += (s, e) => sheetGrid.ClearSelection();

            FlowLayoutPanel selectPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Location = new Point(20, 380),
                Size = new Size(440, 40)
            };

            selectPanel.Controls.Add(btnSelectAll);
            selectPanel.Controls.Add(btnDeselectAll);

            Button btnOk = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Width = 75
            };

            Button btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Width = 75
            };

            FlowLayoutPanel buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                Padding = new Padding(0, 10, 10, 10),
                AutoSize = true
            };

            buttonPanel.Controls.Add(btnCancel);
            buttonPanel.Controls.Add(btnOk);

            Controls.Add(logo);
            Controls.Add(label);
            Controls.Add(sheetGrid);
            Controls.Add(selectPanel);
            Controls.Add(buttonPanel);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SelectedSheets.Clear();
            if (sheetGrid.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow row in sheetGrid.SelectedRows)
                {
                    SelectedSheets.Add(row.Cells[0].Value.ToString());
                }
            }
            base.OnFormClosing(e);
        }
    }
}
