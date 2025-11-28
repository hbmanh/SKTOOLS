using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SKRevitAddins.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Button = System.Windows.Forms.Button;
using Color = System.Drawing.Color;
using ComboBox = System.Windows.Forms.ComboBox;
using Form = System.Windows.Forms.Form;
using Panel = System.Windows.Forms.Panel;
using Point = System.Drawing.Point;
using RadioButton = System.Windows.Forms.RadioButton;
using TextBox = System.Windows.Forms.TextBox;
using View = Autodesk.Revit.DB.View;

namespace SKRevitAddins.GENTools
{
    [Transaction(TransactionMode.Manual)]
    public class RoomDataToElementsCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // --- BƯỚC 1: CHỌN CHẾ ĐỘ (Mode Selection) ---
                bool isManualMode = true;
                using (var modeForm = new ModeSelectionForm())
                {
                    if (modeForm.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;
                    isManualMode = modeForm.IsManualSelection;
                }

                // --- BƯỚC 2: THU THẬP DỮ LIỆU ---
                List<Element> targetElements = new List<Element>();
                List<Room> sourceRooms = new List<Room>();

                // Lấy Phase của View hiện tại để tìm Room chính xác
                ElementId phaseId = doc.ActiveView.get_Parameter(BuiltInParameter.VIEW_PHASE).AsElementId();
                Phase phase = doc.GetElement(phaseId) as Phase;
                if (phase == null)
                {
                    TaskDialog.Show("Error", "Không xác định được Phase của View hiện tại.");
                    return Result.Failed;
                }

                if (isManualMode)
                {
                    // a. Chọn Room nguồn
                    try
                    {
                        IList<Reference> roomRefs = uiDoc.Selection.PickObjects(ObjectType.Element, new RoomFilter(), "Bước 1: Chọn các Room nguồn");
                        foreach (var r in roomRefs)
                        {
                            Room room = doc.GetElement(r) as Room;
                            if (room != null && room.Location != null) sourceRooms.Add(room);
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }

                    if (sourceRooms.Count == 0) return Result.Cancelled;

                    // b. Chọn Target Elements (LOẠI BỎ ROOM RA KHỎI LẦN QUÉT NÀY)
                    try
                    {
                        // TargetElementFilter ở dưới đã được cấu hình để return false với Room
                        IList<Reference> elemRefs = uiDoc.Selection.PickObjects(ObjectType.Element, new TargetElementFilter(), "Bước 2: Quét chọn thiết bị (Đã lọc bỏ Room)");
                        foreach (var r in elemRefs)
                        {
                            Element e = doc.GetElement(r);
                            targetElements.Add(e);
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }
                }
                else // Auto Mode
                {
                    // Lấy tất cả Room trong View
                    sourceRooms = new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .OfClass(typeof(SpatialElement))
                        .Where(e => e is Room)
                        .Cast<Room>()
                        .Where(r => r.Location != null && r.Area > 0)
                        .ToList();

                    // Lấy tất cả Model Elements trong View (trừ Room/Space)
                    targetElements = new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .WhereElementIsNotElementType()
                        .WhereElementIsViewIndependent()
                        .Where(e => e.Category != null && !(e is Room) && !(e is Space) && e.Category.CategoryType == CategoryType.Model)
                        .ToList();
                }

                if (targetElements.Count == 0 || sourceRooms.Count == 0)
                {
                    TaskDialog.Show("Info", "Không tìm thấy dữ liệu phù hợp.");
                    return Result.Succeeded;
                }

                // --- BƯỚC 3: PHÂN TÍCH & CHỌN PARAMETER (ÁP DỤNG CHO CẢ 2 CHẾ ĐỘ) ---
                // Lấy danh sách parameter từ đối tượng đầu tiên tìm được
                List<string> writableParams = GetWritableStringParameters(targetElements.First());

                // Backup: Nếu không tìm thấy, ít nhất cũng thêm các param cơ bản
                if (!writableParams.Contains("Comments")) writableParams.Insert(0, "Comments");
                if (!writableParams.Contains("Mark")) writableParams.Insert(1, "Mark");

                string targetParamName = "";
                using (var paramForm = new ParamSelectionForm(writableParams))
                {
                    if (paramForm.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;
                    targetParamName = paramForm.SelectedParam;
                }

                // --- BƯỚC 4: THỰC THI ---
                int successCount = 0;
                int failCount = 0; // Lỗi do readonly hoặc không tìm thấy param
                int skipCount = 0; // Bỏ qua do không thuộc Room đã chọn (Manual Mode)

                using (Transaction t = new Transaction(doc, "Copy Room Name to Elements"))
                {
                    t.Start();

                    // Tạo HashSet ID Room đã chọn để tra cứu nhanh
                    HashSet<ElementId> selectedRoomIds = new HashSet<ElementId>(sourceRooms.Select(r => r.Id));

                    foreach (Element elem in targetElements)
                    {
                        Parameter p = elem.LookupParameter(targetParamName);
                        // Bỏ qua nếu không có tham số hoặc tham số bị khóa
                        if (p == null || p.IsReadOnly || p.StorageType != StorageType.String)
                        {
                            failCount++;
                            continue;
                        }

                        // Tìm Room chứa đối tượng này
                        Room foundRoom = GetRoomOfElement(elem, phase);

                        if (foundRoom != null)
                        {
                            // LOGIC QUAN TRỌNG:
                            // Nếu Manual Mode: Kiểm tra xem Room tìm thấy có nằm trong danh sách user chọn lúc đầu không?
                            // Nếu không -> Bỏ qua (Skip), không ghi đè giá trị cũ.
                            if (isManualMode && !selectedRoomIds.Contains(foundRoom.Id))
                            {
                                skipCount++;
                                continue;
                            }

                            // Lấy tên Room
                            string roomName = foundRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                            // Có thể lấy thêm Number: string val = $"{foundRoom.Number}: {roomName}";

                            if (p.Set(roomName))
                                successCount++;
                            else
                                failCount++;
                        }
                        else
                        {
                            // Đối tượng không nằm trong Room nào cả
                            failCount++;
                        }
                    }
                    t.Commit();
                }

                string msg = $"Hoàn tất cập nhật tham số '{targetParamName}'.\n\n" +
                             $"- Thành công: {successCount}\n" +
                             $"- Bỏ qua (Không thuộc Room đã chọn): {skipCount}\n" +
                             $"- Thất bại (Không có Room/Param lỗi): {failCount}";

                TaskDialog.Show("Kết quả", msg);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        // --- HELPERS ---

        private List<string> GetWritableStringParameters(Element e)
        {
            var paramsList = new List<string>();
            foreach (Parameter p in e.Parameters)
            {
                if (p.StorageType == StorageType.String && !p.IsReadOnly && p.UserModifiable)
                {
                    paramsList.Add(p.Definition.Name);
                }
            }
            paramsList.Sort();
            return paramsList.Distinct().ToList();
        }

        private Room GetRoomOfElement(Element e, Phase phase)
        {
            if (e is FamilyInstance fi)
            {
                if (fi.Room != null) return fi.Room;
                if (fi.get_Room(phase) != null) return fi.get_Room(phase);
            }
            Location loc = e.Location;
            if (loc is LocationPoint lp)
            {
                return e.Document.GetRoomAtPoint(lp.Point, phase);
            }
            return null;
        }

        // Filter 1: Chỉ cho chọn Room
        public class RoomFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is Room;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        // Filter 2: Cho chọn mọi thứ TRỪ Room (để tránh user chọn nhầm Room làm target)
        public class TargetElementFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                // Logic: Element hợp lệ VÀ KHÔNG PHẢI ROOM, KHÔNG PHẢI SPACE
                return elem.Category != null && !(elem is Room) && !(elem is Space);
            }
            public bool AllowReference(Reference reference, XYZ position) => false;
        }


        // --- UI FORMS ---

        // FORM 1: CHỌN CHẾ ĐỘ
        public class ModeSelectionForm : Form
        {
            private RadioButton rbManual;
            private RadioButton rbAll;
            private Button btnNext;
            private Button btnCancel;

            public bool IsManualSelection => rbManual.Checked;

            public ModeSelectionForm()
            {
                SetupUI();
            }

            private void SetupUI()
            {
                Text = "SKRevit - Select Mode";
                Size = new Size(400, 280);
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false; MinimizeBox = false;
                BackColor = Color.WhiteSmoke;
                Font = SystemFonts.MessageBoxFont;

                var pnlHeader = CreateHeaderPanel();
                pnlHeader.Dock = DockStyle.Top;
                pnlHeader.Height = 45;

                var pnlContent = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
                var gb = new GroupBox { Text = "Selection Mode", Dock = DockStyle.Top, Height = 100 };

                rbManual = new RadioButton { Text = "Option 1: Chọn Room + Quét chọn Đối tượng", Location = new Point(20, 30), AutoSize = true, Checked = true };
                rbAll = new RadioButton { Text = "Option 2: Tự động toàn bộ (Active View)", Location = new Point(20, 60), AutoSize = true };
                gb.Controls.Add(rbManual); gb.Controls.Add(rbAll);
                pnlContent.Controls.Add(gb);

                var pnlBtn = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 50, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(10) };
                btnNext = new Button { Text = "Next >", Width = 90, Height = 30, BackColor = Color.White, DialogResult = DialogResult.OK };
                btnCancel = new Button { Text = "Cancel", Width = 90, Height = 30, BackColor = Color.White, DialogResult = DialogResult.Cancel };
                pnlBtn.Controls.Add(btnNext); pnlBtn.Controls.Add(btnCancel);

                Controls.Add(pnlContent);
                Controls.Add(pnlBtn);
                Controls.Add(pnlHeader);
                AcceptButton = btnNext; CancelButton = btnCancel;
            }
        }

        // FORM 2: CHỌN PARAMETER (HIỆN SAU KHI QUÉT XONG)
        public class ParamSelectionForm : Form
        {
            private ComboBox cbParam;
            private Button btnRun;
            private Button btnCancel;

            public string SelectedParam => cbParam.Text;

            public ParamSelectionForm(List<string> paramsList)
            {
                SetupUI(paramsList);
            }

            private void SetupUI(List<string> paramsList)
            {
                Text = "SKRevit - Select Target Parameter";
                Size = new Size(400, 220);
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false; MinimizeBox = false;
                BackColor = Color.WhiteSmoke;
                Font = SystemFonts.MessageBoxFont;

                var pnlHeader = CreateHeaderPanel();
                pnlHeader.Dock = DockStyle.Top;
                pnlHeader.Height = 45;

                var pnlContent = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
                var lbl = new Label { Text = "Chọn tham số cần điền tên Room vào:", Dock = DockStyle.Top, Height = 25 };

                cbParam = new ComboBox
                {
                    DataSource = paramsList,
                    Dock = DockStyle.Top,
                    DropDownStyle = ComboBoxStyle.DropDown, // Cho phép gõ thêm nếu muốn
                    AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                    AutoCompleteSource = AutoCompleteSource.ListItems
                };

                if (paramsList.Contains("Comments")) cbParam.SelectedItem = "Comments";

                pnlContent.Controls.Add(cbParam);
                pnlContent.Controls.Add(lbl);

                var pnlBtn = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 50, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(10) };
                btnRun = new Button { Text = "Run", Width = 90, Height = 30, BackColor = Color.White, DialogResult = DialogResult.OK };
                btnCancel = new Button { Text = "Cancel", Width = 90, Height = 30, BackColor = Color.White, DialogResult = DialogResult.Cancel };
                pnlBtn.Controls.Add(btnRun); pnlBtn.Controls.Add(btnCancel);

                Controls.Add(pnlContent);
                Controls.Add(pnlBtn);
                Controls.Add(pnlHeader);
                AcceptButton = btnRun; CancelButton = btnCancel;
            }
        }

        private static Panel CreateHeaderPanel()
        {
            var headerPanel = new Panel { BackColor = Color.White };
            var logoBox = new PictureBox { SizeMode = PictureBoxSizeMode.Zoom, Size = new Size(30, 30), Location = new Point(10, 7) };
            try { LogoHelper.TryLoadLogo(logoBox); } catch { }
            headerPanel.Controls.Add(logoBox);
            var companyLabel = new Label { Text = "Shinken Group®", Font = new Font("Segoe UI", 9F, FontStyle.Bold), AutoSize = true, Location = new Point(50, 12), ForeColor = Color.Black };
            headerPanel.Controls.Add(companyLabel);
            return headerPanel;
        }
    }
}