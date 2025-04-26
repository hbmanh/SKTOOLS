using System;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Form = System.Windows.Forms.Form;
using Panel = System.Windows.Forms.Panel;
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

            // Hộp thoại nhập TEXT_FONT
            string userInputFontName = "";

            using (var form = new Form())
            {
                form.Size = new System.Drawing.Size(340, 220);
                form.Text = "Shinken Group® - Text Note Editor";
                form.FormBorderStyle = FormBorderStyle.FixedSingle;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                // Tạo panel nhỏ để chứa logo và text
                var headerPanel = new Panel();
                headerPanel.Size = new System.Drawing.Size(form.ClientSize.Width, 40);
                headerPanel.Location = new System.Drawing.Point(0, 0);
                headerPanel.BackColor = System.Drawing.Color.White;
                form.Controls.Add(headerPanel);

                try
                {
                    var logoPictureBox = new PictureBox();
                    logoPictureBox.Image = System.Drawing.Image.FromFile("C:\\ProgramData\\Autodesk\\Revit\\Addins\\2023\\SKTools.bundle\\Contents\\Resources\\Images\\shinken.png"); // đường dẫn logo mới
                    logoPictureBox.SizeMode = PictureBoxSizeMode.Zoom;
                    logoPictureBox.Size = new System.Drawing.Size(30, 30);
                    logoPictureBox.Location = new System.Drawing.Point(10, 5);
                    headerPanel.Controls.Add(logoPictureBox);

                    var companyLabel = new Label();
                    companyLabel.Text = "Shinken Group®";
                    companyLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
                    companyLabel.AutoSize = true;
                    companyLabel.Location = new System.Drawing.Point(50, 10); // Bên phải logo
                    headerPanel.Controls.Add(companyLabel);
                }
                catch (System.IO.FileNotFoundException)
                {
                    // Không tìm thấy hình
                }

                // Label "Chọn Font"
                var fontLabel = new Label();
                fontLabel.Text = "Chọn Font:";
                fontLabel.Location = new System.Drawing.Point(20, 55);
                fontLabel.Size = new System.Drawing.Size(280, 20);
                fontLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
                form.Controls.Add(fontLabel);

                // Textbox Font
                var fontTextBox = new TextBox();
                fontTextBox.Location = new System.Drawing.Point(20, 80);
                fontTextBox.Size = new System.Drawing.Size(300, 25);
                fontTextBox.Text = "游ゴシック";
                form.Controls.Add(fontTextBox);

                // Info Label
                var infoLabel = new Label();
                infoLabel.Text = "※ Chiều rộng text sẽ tự động điều chỉnh.";
                infoLabel.Location = new System.Drawing.Point(20, 110);
                infoLabel.Size = new System.Drawing.Size(300, 20);
                infoLabel.ForeColor = System.Drawing.Color.Gray;
                infoLabel.Font = new System.Drawing.Font("Segoe UI", 8F);
                form.Controls.Add(infoLabel);

                // OK Button
                var okButton = new Button();
                okButton.Text = "OK";
                okButton.Size = new System.Drawing.Size(90, 30);
                okButton.Location = new System.Drawing.Point(form.ClientSize.Width / 2 - 100, 140);
                okButton.Click += (sender, e) =>
                {
                    userInputFontName = fontTextBox.Text;
                    form.DialogResult = DialogResult.OK;
                    form.Close();
                };
                form.Controls.Add(okButton);

                // Cancel Button
                var cancelButton = new Button();
                cancelButton.Text = "Cancel";
                cancelButton.Size = new System.Drawing.Size(90, 30);
                cancelButton.Location = new System.Drawing.Point(form.ClientSize.Width / 2 + 10, 140);
                cancelButton.Click += (sender, e) =>
                {
                    form.DialogResult = DialogResult.Cancel;
                    form.Close();
                };
                form.Controls.Add(cancelButton);

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

                        string textNoteTypeName = $"AWS-{fontName} {textSize} ({widthFactor}) ({GetColorRGB(color)})";
                        string baseName = textNoteTypeName;

                        var existingType = new FilteredElementCollector(doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .FirstOrDefault(t => t.Name.Equals(textNoteTypeName));
                        textNoteTypeName = $"{baseName}";

                        var exitsTextNoteTypeName = new FilteredElementCollector(doc)
                            .OfClass(typeof(TextNoteType))
                            .FirstOrDefault(t => t.Name.Equals(textNoteTypeName))?.Id;

                        if (exitsTextNoteTypeName != null)
                        {
                            TextNoteType exitsTextNoteType = doc.GetElement(exitsTextNoteTypeName) as TextNoteType;
                            // Đổi Type của các Text Notes trong nhóm
                            foreach (var textNote in group)
                            {
                                textNote.ChangeTypeId(exitsTextNoteType.Id);
                                textNote.get_Parameter(BuiltInParameter.KEEP_READABLE).Set(1);

                                // Tính toán chiều rộng phù hợp với nội dung
                                double minWidth = textNote.GetMinimumAllowedWidth();
                                double maxWidth = textNote.GetMaximumAllowedWidth();

                                // Tính toán chiều rộng dựa trên nội dung của TextNote
                                string content = textNote.Text;
                                // Ước tính chiều rộng dựa trên độ dài nội dung và kích thước chữ
                                // Hệ số 0.85 thay vì 0.7 để tăng thêm khoảng an toàn
                                double estimatedWidth = content.Length * (textSize / 304.8) * widthFactor * 0.85;

                                // Thêm 10% vào chiều rộng để tránh sai sót do font chữ
                                estimatedWidth = estimatedWidth * 1.1;

                                // Đảm bảo chiều rộng nằm trong phạm vi cho phép
                                double adjustedWidth = Math.Max(minWidth, Math.Min(estimatedWidth, maxWidth));

                                // Sử dụng chiều rộng tính toán
                                textNote.Width = adjustedWidth;
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

                                // Tính toán chiều rộng phù hợp với nội dung
                                double minWidth = textNote.GetMinimumAllowedWidth();
                                double maxWidth = textNote.GetMaximumAllowedWidth();

                                // Tính toán chiều rộng dựa trên nội dung của TextNote
                                string content = textNote.Text;
                                // Ước tính chiều rộng dựa trên độ dài nội dung và kích thước chữ
                                // Hệ số 0.85 thay vì 0.7 để tăng thêm khoảng an toàn
                                double estimatedWidth = content.Length * (textSize / 304.8) * widthFactor * 0.85;

                                // Thêm 10% vào chiều rộng để tránh sai sót do font chữ
                                estimatedWidth = estimatedWidth * 1.1;

                                // Đảm bảo chiều rộng nằm trong phạm vi cho phép
                                double adjustedWidth = Math.Max(minWidth, Math.Min(estimatedWidth, maxWidth));

                                // Sử dụng chiều rộng tính toán
                                textNote.Width = adjustedWidth;
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