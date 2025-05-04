using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Form = System.Windows.Forms.Form;
using View = Autodesk.Revit.DB.View;
using Point = System.Drawing.Point;
using Color = System.Drawing.Color;

namespace SKRevitAddins.GridTools
{
    public class GridControlForm : Form
    {
        private readonly UIDocument uidoc;
        private readonly Document doc;
        private Button btnToggleMode;
        private Button btnToggleBubbles;
        private Button btnHideBubbles; // Nút mới để tắt bubble

        public GridControlForm(UIDocument uidoc)
        {
            this.uidoc = uidoc;
            this.doc = uidoc.Document;

            Text = "Shinken Group® - Grid Control";
            Size = new Size(320, 250); // Tăng chiều cao form để chứa nút mới
            StartPosition = FormStartPosition.CenterScreen;

            PictureBox logo = new PictureBox
            {
                Image = Image.FromFile("C:\\ProgramData\\Autodesk\\Revit\\Addins\\2023\\SKTools.bundle\\Contents\\Resources\\Images\\shinken.png"),
                Size = new Size(32, 32),
                Location = new Point(10, 10),
                SizeMode = PictureBoxSizeMode.StretchImage
            };

            Label lblTitle = new Label
            {
                Text = "   Shinken Group®",
                Location = new Point(50, 15),
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            btnToggleMode = new Button
            {
                Text = "Chuyển thành Grid 2D",
                Size = new Size(240, 35),
                Location = new Point(30, 60)
            };
            btnToggleMode.Click += Toggle2D3D_Click;

            btnToggleBubbles = new Button
            {
                Text = "Đổi đầu Bubble (Selected)",
                Size = new Size(240, 35),
                Location = new Point(30, 105)
            };
            btnToggleBubbles.Click += ToggleBubbles_Click;

            // Nút mới để tắt/bật bubble
            btnHideBubbles = new Button
            {
                Text = "Tắt/Hiện Bubble (Selected)",
                Size = new Size(240, 35),
                Location = new Point(30, 150)
            };
            btnHideBubbles.Click += ToggleHideBubbles_Click;

            Controls.Add(logo);
            Controls.Add(lblTitle);
            Controls.Add(btnToggleMode);
            Controls.Add(btnToggleBubbles);
            Controls.Add(btnHideBubbles); // Thêm nút mới vào form
        }

        private void Toggle2D3D_Click(object sender, EventArgs e)
        {
            View activeView = doc.ActiveView;

            var grids = new FilteredElementCollector(doc, activeView.Id)
                .WhereElementIsNotElementType()
                .OfClass(typeof(Grid))
                .OfCategory(BuiltInCategory.OST_Grids)
                .Cast<Grid>()
                .ToList();

            if (!grids.Any())
            {
                TaskDialog.Show("Thông báo", "Không tìm thấy Grid trong View.");
                return;
            }

            // Chuyển tất cả grid sang 2D (View Specific)
            using (Transaction t = new Transaction(doc, "Chuyển Grid sang 2D"))
            {
                t.Start();
                foreach (var grid in grids)
                {
                    grid.SetDatumExtentType(DatumEnds.End0, activeView, DatumExtentType.ViewSpecific);
                    grid.SetDatumExtentType(DatumEnds.End1, activeView, DatumExtentType.ViewSpecific);
                }
                t.Commit();
            }

            TaskDialog.Show("Xong", "Đã chuyển tất cả Grid sang chế độ 2D.");
        }

        private void ToggleBubbles_Click(object sender, EventArgs e)
        {
            View activeView = doc.ActiveView;

            var selectedIds = uidoc.Selection.GetElementIds();
            if (!selectedIds.Any())
            {
                TaskDialog.Show("Thông báo", "Vui lòng chọn ít nhất một Grid trong View.");
                return;
            }

            var allGrids = new FilteredElementCollector(doc, activeView.Id)
                .WhereElementIsNotElementType()
                .OfClass(typeof(Grid))
                .OfCategory(BuiltInCategory.OST_Grids)
                .Cast<Grid>()
                .ToList();

            var selectedGrids = selectedIds.Select(id => doc.GetElement(id))
                                           .Where(el => el is Grid)
                                           .Cast<Grid>()
                                           .ToList();

            if (!selectedGrids.Any())
            {
                TaskDialog.Show("Thông báo", "Các đối tượng được chọn không phải là Grid.");
                return;
            }

            using (Transaction t = new Transaction(doc, "Chuyển Bubble đầu/cuối"))
            {
                t.Start();

                // Đưa toàn bộ grid về chế độ 2D
                foreach (var grid in allGrids)
                {
                    grid.SetDatumExtentType(DatumEnds.End0, activeView, DatumExtentType.ViewSpecific);
                    grid.SetDatumExtentType(DatumEnds.End1, activeView, DatumExtentType.ViewSpecific);
                }

                // Chỉ đổi đầu bubble với các grid được chọn
                foreach (var grid in selectedGrids)
                {
                    bool bubbleEnd0 = grid.IsBubbleVisibleInView(DatumEnds.End0, activeView);
                    bool bubbleEnd1 = grid.IsBubbleVisibleInView(DatumEnds.End1, activeView);

                    if (bubbleEnd0)
                    {
                        grid.HideBubbleInView(DatumEnds.End0, activeView);
                        grid.ShowBubbleInView(DatumEnds.End1, activeView);
                    }
                    else
                    {
                        grid.HideBubbleInView(DatumEnds.End1, activeView);
                        grid.ShowBubbleInView(DatumEnds.End0, activeView);
                    }
                }
                t.Commit();
            }

            TaskDialog.Show("Xong", "Đã chuyển Bubble đầu/cuối cho các Grid được chọn.");
        }

        // Phương thức mới để tắt/hiện bubble của grid được chọn
        private void ToggleHideBubbles_Click(object sender, EventArgs e)
        {
            View activeView = doc.ActiveView;

            var selectedIds = uidoc.Selection.GetElementIds();
            if (!selectedIds.Any())
            {
                TaskDialog.Show("Thông báo", "Vui lòng chọn ít nhất một Grid trong View.");
                return;
            }

            var selectedGrids = selectedIds.Select(id => doc.GetElement(id))
                                          .Where(el => el is Grid)
                                          .Cast<Grid>()
                                          .ToList();

            if (!selectedGrids.Any())
            {
                TaskDialog.Show("Thông báo", "Các đối tượng được chọn không phải là Grid.");
                return;
            }

            using (Transaction t = new Transaction(doc, "Tắt/Hiện Bubble Grid"))
            {
                t.Start();

                foreach (var grid in selectedGrids)
                {
                    // Kiểm tra trạng thái hiện tại của bubble
                    bool bubbleEnd0 = grid.IsBubbleVisibleInView(DatumEnds.End0, activeView);
                    bool bubbleEnd1 = grid.IsBubbleVisibleInView(DatumEnds.End1, activeView);

                    // Nếu ít nhất một đầu có bubble, sẽ ẩn tất cả
                    if (bubbleEnd0 || bubbleEnd1)
                    {
                        // Ẩn cả hai đầu bubble
                        grid.HideBubbleInView(DatumEnds.End0, activeView);
                        grid.HideBubbleInView(DatumEnds.End1, activeView);
                    }
                    else
                    {
                        // Hiện một đầu bubble nếu cả hai đang ẩn (mặc định hiện đầu thứ nhất)
                        grid.ShowBubbleInView(DatumEnds.End0, activeView);
                    }
                }

                t.Commit();
            }

            TaskDialog.Show("Xong", "Đã tắt/hiện Bubble cho các Grid được chọn.");
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class GridToolsCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Application.EnableVisualStyles();
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            GridControlForm form = new GridControlForm(uidoc);
            form.ShowDialog();
            return Result.Succeeded;
        }
    }
}