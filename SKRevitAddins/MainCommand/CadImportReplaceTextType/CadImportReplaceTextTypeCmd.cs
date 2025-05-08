using System;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Form = System.Windows.Forms.Form;
using Panel = System.Windows.Forms.Panel;
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

            var textNotes = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfClass(typeof(TextNote))
                .Cast<TextNote>()
                .ToList();

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

            using (var form = new Form())
            {
                form.Size = new System.Drawing.Size(340, 220);
                form.Text = "Shinken Group® - Text Note Editor";
                form.FormBorderStyle = FormBorderStyle.FixedSingle;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var headerPanel = new Panel
                {
                    Size = new System.Drawing.Size(form.ClientSize.Width, 40),
                    Location = new System.Drawing.Point(0, 0),
                    BackColor = System.Drawing.Color.White
                };
                form.Controls.Add(headerPanel);

                try
                {
                    var logoPictureBox = new PictureBox
                    {
                        Image = System.Drawing.Image.FromFile("C:\\ProgramData\\Autodesk\\Revit\\Addins\\2023\\SKTools.bundle\\Contents\\Resources\\Images\\shinken.png"),
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Size = new System.Drawing.Size(30, 30),
                        Location = new System.Drawing.Point(10, 5)
                    };
                    headerPanel.Controls.Add(logoPictureBox);

                    var companyLabel = new Label
                    {
                        Text = "Shinken Group®",
                        Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold),
                        AutoSize = true,
                        Location = new System.Drawing.Point(50, 10)
                    };
                    headerPanel.Controls.Add(companyLabel);
                }
                catch { }

                var fontLabel = new Label
                {
                    Text = "Chọn Font:",
                    Location = new System.Drawing.Point(20, 55),
                    Size = new System.Drawing.Size(280, 20),
                    Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold)
                };
                form.Controls.Add(fontLabel);

                var fontTextBox = new TextBox
                {
                    Location = new System.Drawing.Point(20, 80),
                    Size = new System.Drawing.Size(300, 25),
                    Text = "游ゴシック"
                };
                form.Controls.Add(fontTextBox);

                var infoLabel = new Label
                {
                    Text = "※ Chiều rộng text sẽ tự động điều chỉnh.",
                    Location = new System.Drawing.Point(20, 110),
                    Size = new System.Drawing.Size(300, 20),
                    ForeColor = System.Drawing.Color.Gray,
                    Font = new System.Drawing.Font("Segoe UI", 8F)
                };
                form.Controls.Add(infoLabel);

                var okButton = new Button
                {
                    Text = "OK",
                    Size = new System.Drawing.Size(90, 30),
                    Location = new System.Drawing.Point(form.ClientSize.Width / 2 - 100, 140)
                };
                okButton.Click += (sender, e) =>
                {
                    userInputFontName = fontTextBox.Text;
                    form.DialogResult = DialogResult.OK;
                    form.Close();
                };
                form.Controls.Add(okButton);

                var cancelButton = new Button
                {
                    Text = "Cancel",
                    Size = new System.Drawing.Size(90, 30),
                    Location = new System.Drawing.Point(form.ClientSize.Width / 2 + 10, 140)
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

            if (groupedTextNotes.Any())
            {
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

                        if (existingType != null)
                        {
                            foreach (var textNote in group)
                            {
                                textNote.ChangeTypeId(existingType.Id);
                                textNote.get_Parameter(BuiltInParameter.KEEP_READABLE).Set(1);
                                string content = textNote.Text;
                                double estimatedWidth = GetEstimatedWidth(content, textSize, widthFactor);
                                double minWidth = textNote.GetMinimumAllowedWidth();
                                double maxWidth = textNote.GetMaximumAllowedWidth();
                                textNote.Width = Math.Max(minWidth, Math.Min(estimatedWidth, maxWidth));
                            }
                        }
                        else
                        {
                            var defaultTypeId = new FilteredElementCollector(doc)
                                .OfClass(typeof(TextNoteType))
                                .Cast<TextNoteType>()
                                .FirstOrDefault()?.Id;
                            var newTypeId = (doc.GetElement(defaultTypeId) as TextNoteType)?.Duplicate(typeName)?.Id;
                            var newType = doc.GetElement(newTypeId) as TextNoteType;

                            if (newType != null)
                            {
                                newType.get_Parameter(BuiltInParameter.TEXT_FONT).Set(fontName);
                                newType.get_Parameter(BuiltInParameter.TEXT_SIZE).Set(textSize / 304.8);
                                newType.get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE).Set(widthFactor);
                                newType.get_Parameter(BuiltInParameter.TEXT_BACKGROUND).Set(1);
                                newType.get_Parameter(BuiltInParameter.LINE_COLOR).Set(color);
                            }

                            foreach (var textNote in group)
                            {
                                textNote.ChangeTypeId(newType.Id);
                                textNote.get_Parameter(BuiltInParameter.KEEP_READABLE).Set(1);
                                string content = textNote.Text;
                                double estimatedWidth = GetEstimatedWidth(content, textSize, widthFactor);
                                double minWidth = textNote.GetMinimumAllowedWidth();
                                double maxWidth = textNote.GetMaximumAllowedWidth();
                                textNote.Width = Math.Max(minWidth, Math.Min(estimatedWidth, maxWidth));
                            }
                        }
                    }
                    trans.Commit();
                    TaskDialog.Show("Success", "Successfully applied new TextType with auto sizing.");
                }
            }

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
            double totalWeight = 0;
            foreach (char c in content)
            {
                totalWeight += IsCJK(c) ? 1.8 : 1.0;
            }
            double estimatedWidth = totalWeight * (textSize / 304.8) * widthFactor;
            return estimatedWidth * 1.1;
        }

        private bool IsCJK(char c)
        {
            return (c >= 0x3000 && c <= 0x9FFF) || (c >= 0xFF00 && c <= 0xFFEF);
        }
    }
}
