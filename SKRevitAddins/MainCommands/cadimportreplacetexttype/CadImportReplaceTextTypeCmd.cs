using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Color = System.Drawing.Color;
using Form = System.Windows.Forms.Form;
using Panel = System.Windows.Forms.Panel;
using Point = System.Drawing.Point;
using TextBox = System.Windows.Forms.TextBox;

namespace SKRevitAddins.CadImportReplaceTextType
{
    [Transaction(TransactionMode.Manual)]
    public class CadImportReplaceTextTypeCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Lấy TextNote từ selection hoặc toàn bộ view
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            List<TextNote> textNotes;

            if (selectedIds.Count == 0)
            {
                textNotes = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfClass(typeof(TextNote))
                    .Cast<TextNote>()
                    .ToList();
            }
            else
            {
                textNotes = selectedIds
                    .Select(id => doc.GetElement(id))
                    .OfType<TextNote>()
                    .ToList();
            }

            if (!textNotes.Any())
            {
                TaskDialog.Show("Thông báo", "Không tìm thấy TextNote.");
                return Result.Cancelled;
            }

            var groupedTextNotes = textNotes
                .GroupBy(tn => new
                {
                    FontName = tn.Symbol.get_Parameter(BuiltInParameter.TEXT_FONT)?.AsString(),
                    TextSize = Math.Round(doc.GetElement(tn.GetTypeId()).get_Parameter(BuiltInParameter.TEXT_SIZE).AsDouble() * 304.8, 1),
                    WidthFactor = Math.Round(doc.GetElement(tn.GetTypeId()).get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE).AsDouble(), 1),
                    Color = GetTextColor(doc, tn)
                })
                .ToList();

            string userInputFontName = "";
            bool autoCalculateWidth = true;
            double fixedWidth = 0;

            using (var form = new Form())
            {
                form.Size = new Size(340, 280);
                form.Text = "Shinken Group® - Text Note Editor";
                form.FormBorderStyle = FormBorderStyle.FixedSingle;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var headerPanel = new Panel
                {
                    Size = new Size(form.ClientSize.Width, 40),
                    Location = new Point(0, 0),
                    BackColor = Color.White
                };
                form.Controls.Add(headerPanel);

                try
                {
                    string logoPath = Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        "SKTools.bundle", "Icon", "shinken.png");

                    if (File.Exists(logoPath))
                    {
                        var logoPictureBox = new PictureBox
                        {
                            SizeMode = PictureBoxSizeMode.Zoom,
                            Size = new Size(30, 30),
                            Location = new Point(10, 5),
                            Image = Image.FromFile(logoPath)
                        };
                        headerPanel.Controls.Add(logoPictureBox);
                    }

                    var companyLabel = new Label
                    {
                        Text = "Shinken Group®",
                        Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                        AutoSize = true,
                        Location = new Point(50, 10)
                    };
                    headerPanel.Controls.Add(companyLabel);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error loading logo: " + ex.Message);
                }


                var fontLabel = new Label
                {
                    Text = "Chọn Font:",
                    Location = new Point(20, 50),
                    Size = new Size(280, 20),
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold)
                };
                form.Controls.Add(fontLabel);

                var fontTextBox = new TextBox
                {
                    Location = new Point(20, 75),
                    Size = new Size(300, 25),
                    Text = "游ゴシック"
                };
                form.Controls.Add(fontTextBox);

                var autoWidthCheckbox = new CheckBox
                {
                    Text = "Tự động tính chiều rộng",
                    Checked = true,
                    Location = new Point(20, 110),
                    Size = new Size(300, 20)
                };
                form.Controls.Add(autoWidthCheckbox);

                var widthLabel = new Label
                {
                    Text = "Chiều rộng cố định (nếu không auto):",
                    Location = new Point(20, 135),
                    Size = new Size(300, 20)
                };
                form.Controls.Add(widthLabel);

                var widthTextBox = new TextBox
                {
                    Location = new Point(20, 160),
                    Size = new Size(300, 25),
                    Enabled = false
                };
                form.Controls.Add(widthTextBox);

                autoWidthCheckbox.CheckedChanged += (s, e) =>
                {
                    widthTextBox.Enabled = !autoWidthCheckbox.Checked;
                };

                var okButton = new Button
                {
                    Text = "OK",
                    Size = new Size(90, 30),
                    Location = new Point(form.ClientSize.Width / 2 - 100, 200)
                };
                okButton.Click += (sender, e) =>
                {
                    userInputFontName = fontTextBox.Text;
                    autoCalculateWidth = autoWidthCheckbox.Checked;
                    if (!autoCalculateWidth)
                    {
                        double.TryParse(widthTextBox.Text, out fixedWidth);
                    }
                    form.DialogResult = DialogResult.OK;
                    form.Close();
                };
                form.Controls.Add(okButton);

                var cancelButton = new Button
                {
                    Text = "Cancel",
                    Size = new Size(90, 30),
                    Location = new Point(form.ClientSize.Width / 2 + 10, 200)
                };
                cancelButton.Click += (sender, e) =>
                {
                    form.DialogResult = DialogResult.Cancel;
                    form.Close();
                };
                form.Controls.Add(cancelButton);

                if (form.ShowDialog() == DialogResult.Cancel)
                    return Result.Cancelled;
            }

            using (Transaction trans = new Transaction(doc, "Replace Text Notes"))
            {
                trans.Start();
                foreach (var group in groupedTextNotes)
                {
                    double textSize = group.Key.TextSize;
                    double widthFactor = group.Key.WidthFactor;
                    int color = group.Key.Color;
                    string fontName = string.IsNullOrEmpty(userInputFontName) ? group.Key.FontName : userInputFontName;

                    string typeName = $"AWS-{fontName} {textSize} ({widthFactor}) ({GetColorRGB(color)})";

                    var existingType = new FilteredElementCollector(doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .FirstOrDefault(t => t.Name.Equals(typeName));

                    TextNoteType typeToUse;

                    if (existingType != null)
                    {
                        typeToUse = existingType;
                    }
                    else
                    {
                        var defaultType = new FilteredElementCollector(doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .FirstOrDefault();
                        var newTypeId = defaultType?.Duplicate(typeName)?.Id;
                        typeToUse = doc.GetElement(newTypeId) as TextNoteType;

                        if (typeToUse != null)
                        {
                            typeToUse.get_Parameter(BuiltInParameter.TEXT_FONT).Set(fontName);
                            typeToUse.get_Parameter(BuiltInParameter.TEXT_SIZE).Set(textSize / 304.8);
                            typeToUse.get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE).Set(widthFactor);
                            typeToUse.get_Parameter(BuiltInParameter.TEXT_BACKGROUND).Set(1);
                            typeToUse.get_Parameter(BuiltInParameter.LINE_COLOR).Set(color);
                        }
                    }

                    foreach (var textNote in group)
                    {
                        textNote.ChangeTypeId(typeToUse.Id);
                        textNote.get_Parameter(BuiltInParameter.KEEP_READABLE).Set(1);
                        string content = textNote.Text;

                        double minWidth = textNote.GetMinimumAllowedWidth();
                        double maxWidth = textNote.GetMaximumAllowedWidth();
                        double finalWidth = autoCalculateWidth
                            ? GetEstimatedWidth(content, textSize, widthFactor)
                            : fixedWidth;

                        textNote.Width = Math.Max(minWidth, Math.Min(finalWidth, maxWidth));
                    }
                }
                trans.Commit();
            }

            TaskDialog.Show("Success", "Successfully applied new TextType with auto sizing.");
            return Result.Succeeded;
        }

        private int GetTextColor(Document doc, TextNote textNote)
        {
            return textNote.Symbol.get_Parameter(BuiltInParameter.LINE_COLOR).AsInteger();
        }

        private string GetColorRGB(int color)
        {
            int red = color & 0xFF;
            int green = (color >> 8) & 0xFF;
            int blue = (color >> 16) & 0xFF;
            return $"{red},{green},{blue}";
        }

        private double GetEstimatedWidth(string content, double textSize, double widthFactor)
        {
            // Cắt nội dung theo từng dòng
            var lines = content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            // Duyệt từng dòng để tính "trọng số" dựa trên CJK hoặc Latin
            double maxLineWeight = lines.Max(line =>
            {
                double weight = 0;
                foreach (char c in line)
                {
                    weight += IsCJK(c) ? 1.8 : 1.0; // CJK chiếm rộng hơn
                }
                return weight;
            });

            // Tính width dựa trên dòng dài nhất
            double estimatedWidth = maxLineWeight * (textSize / 304.8) * widthFactor;

            return estimatedWidth * 1.1; // Thêm hệ số đệm
        }


        private bool IsCJK(char c)
        {
            return (c >= 0x3000 && c <= 0x9FFF) || (c >= 0xFF00 && c <= 0xFFEF);
        }
    }
}
