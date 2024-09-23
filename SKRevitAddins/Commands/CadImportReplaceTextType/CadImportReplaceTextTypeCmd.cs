using System;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Form = System.Windows.Forms.Form;
using TextBox = System.Windows.Forms.TextBox;

namespace SKRevitAddins.Commands.CadImportReplaceTextType
{
    [Transaction(TransactionMode.Manual)]
    public class CadImportReplaceTextTypeCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Bước nhóm cho Text Size và Width Factor
            double sizeGroupStep = 0.1;
            double widthFactorStep = 0.1;

            // Lấy ra tất cả Text Notes trong active view
            var textNotes = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfClass(typeof(TextNote))
                .Cast<TextNote>()
                .ToList();

            // Nhóm Text Notes theo Text Size và Width Factor
            var groupedTextNotes = textNotes
                .GroupBy(tn => new
                {
                    FontName = tn.Symbol.get_Parameter(BuiltInParameter.TEXT_FONT)?.AsString(),
                    TextSize = Math.Round(doc.GetElement(tn.GetTypeId()).get_Parameter(BuiltInParameter.TEXT_SIZE).AsDouble() * 304.8, 1),
                    WidthFactor = Math.Round(doc.GetElement(tn.GetTypeId()).get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE).AsDouble(), 1),
                    Color = GetTextColor(doc, tn)
                })
                .ToList();

            // Hộp thoại nhập TEXT_FONT và WIDTH
            string userInputFontName = "";
            double userInputWidth = 100; // Default width

            using (var form = new Form())
            {
                form.Text = "CHỌN TEXT FONT VÀ WIDTH:";
                form.Size = new System.Drawing.Size(300, 200);
                form.FormBorderStyle = FormBorderStyle.FixedSingle;
                form.MaximizeBox = false;

                var fontLabel = new Label();
                fontLabel.Text = "Font:";
                fontLabel.Dock = DockStyle.Top;

                var fontTextBox = new TextBox();
                fontTextBox.Dock = DockStyle.Top;
                fontTextBox.Text = "游ゴシック"; // Default font

                var widthLabel = new Label();
                widthLabel.Text = "Width (Default = 100):";
                widthLabel.Dock = DockStyle.Top;

                var widthTextBox = new TextBox();
                widthTextBox.Dock = DockStyle.Top;

                var cancelButton = new Button();
                cancelButton.Text = "Cancel";
                cancelButton.Width = 80;
                cancelButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
                cancelButton.Click += (sender, e) =>
                {
                    form.DialogResult = DialogResult.Cancel;
                    form.Close();
                };

                var okButton = new Button();
                okButton.Text = "OK";
                okButton.Width = 80;
                okButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
                okButton.Click += (sender, e) =>
                {
                    userInputFontName = fontTextBox.Text;
                    double.TryParse(widthTextBox.Text, out userInputWidth);
                    form.DialogResult = DialogResult.OK;
                    form.Close();
                };

                form.Controls.Add(okButton);
                form.Controls.Add(cancelButton);
                form.Controls.Add(widthTextBox);
                form.Controls.Add(widthLabel);
                form.Controls.Add(fontTextBox);
                form.Controls.Add(fontLabel);

                // Chỉnh vị trí của các controls
                fontLabel.Location = new System.Drawing.Point(10, 10);
                fontTextBox.Location = new System.Drawing.Point(10, 30);
                widthLabel.Location = new System.Drawing.Point(10, 60);
                widthTextBox.Location = new System.Drawing.Point(10, 80);
                cancelButton.Location = new System.Drawing.Point(form.ClientSize.Width - cancelButton.Width - 10, form.ClientSize.Height - cancelButton.Height - 10);
                okButton.Location = new System.Drawing.Point(form.ClientSize.Width - cancelButton.Width - okButton.Width - 20, form.ClientSize.Height - okButton.Height - 10);

                if (form.ShowDialog() == DialogResult.Cancel)
                {
                    return Result.Cancelled;
                }
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
                        int color = group.Key.Color; // Màu của nhóm Text Notes
                        string fontName = string.IsNullOrEmpty(userInputFontName) ? group.Key.FontName : userInputFontName;
                        double width = userInputWidth == 0 ? 100 : userInputWidth;

                        string textNoteTypeName = $"AWS-{fontName} {textSize} ({widthFactor}) ({GetColorRGB(color)})";
                        string baseName = textNoteTypeName;
                        int suffix = 1;

                        //while (true)
                        //{
                        //    var existingType = new FilteredElementCollector(doc)
                        //        .OfClass(typeof(TextNoteType))
                        //        .Cast<TextNoteType>()
                        //        .FirstOrDefault(t => t.Name.Equals(textNoteTypeName));

                        //    if (existingType == null)
                        //    {
                        //        break;
                        //    }

                        //    textNoteTypeName = $"{baseName}_{suffix}";
                        //    suffix++;
                        //}

                        var existingType = new FilteredElementCollector(doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .FirstOrDefault(t => t.Name.Equals(textNoteTypeName));
                        textNoteTypeName = $"{baseName}";

                        var exitsTextNoteTypeName = new FilteredElementCollector(doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .FirstOrDefault(t => t.Name.Equals(textNoteTypeName))?.Id;

                        if (exitsTextNoteTypeName != null)
                        {
                            TextNoteType exitsTextNoteType = doc.GetElement(exitsTextNoteTypeName) as TextNoteType;
                            // Đổi Type của các Text Notes trong nhóm
                            foreach (var textNote in group)
                            {
                                textNote.ChangeTypeId(exitsTextNoteType.Id);
                                textNote.get_Parameter(BuiltInParameter.KEEP_READABLE).Set(1);
                                textNote.Width = width / 304.8;
                            }
                        }
                        else
                        {
                            // Tạo Type mới nếu chưa tồn tại
                            var defautTextNoteType = new FilteredElementCollector(doc)
                                .OfClass(typeof(TextNoteType))
                                .Cast<TextNoteType>()
                                .FirstOrDefault()?.Id; ;
                            var newTypeId = (doc.GetElement(defautTextNoteType) as TextNoteType)?.Duplicate(textNoteTypeName)?.Id;
                            var newType = doc.GetElement(newTypeId) as TextNoteType;
                            if (newType != null)
                            {
                                // Set font, textSize, widthFactor và màu cho Type mới
                                newType.get_Parameter(BuiltInParameter.TEXT_FONT).Set(fontName);
                                newType.get_Parameter(BuiltInParameter.TEXT_SIZE).Set(textSize / 304.8);
                                newType.get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE).Set(widthFactor);
                                newType.get_Parameter(BuiltInParameter.TEXT_BACKGROUND).Set(1);

                                // Set màu cho Type mới
                                newType.get_Parameter(BuiltInParameter.LINE_COLOR).Set(color);
                            }

                            foreach (var textNote in group)
                            {
                                textNote.ChangeTypeId(newType.Id);
                                textNote.get_Parameter(BuiltInParameter.KEEP_READABLE).Set(1);
                                textNote.Width = width / 304.8;
                            }
                        }
                    }
                    trans.Commit();
                    TaskDialog.Show("Success", "Successfully applied new TextType.");
                }
            }

            return Result.Succeeded;
        }

        private int GetTextColor(Document doc, TextNote textNote)
        {
            // Lấy giá trị màu của TextNote dựa vào BuiltInParameter.LINE_COLOR
            Parameter param = textNote.Symbol.get_Parameter(BuiltInParameter.LINE_COLOR);
            int color = param.AsInteger();
            return color;
        }

        private string GetColorRGB(int color)
        {
            // Chuyển đổi mã màu từ giá trị RGB sang chuỗi
            int red = color & 0xFF;
            int green = (color >> 8) & 0xFF;
            int blue = (color >> 16) & 0xFF;
            return $"{red},{green},{blue}";
        }
    }
}
