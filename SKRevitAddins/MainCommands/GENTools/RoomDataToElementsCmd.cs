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
using CheckBox = System.Windows.Forms.CheckBox;
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
                // --- BƯỚC 1: CHỌN CHẾ ĐỘ ---
                bool isManualMode = true;
                bool includeLinks = false; // Tùy chọn mới

                using (var modeForm = new ModeSelectionForm())
                {
                    if (modeForm.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;
                    isManualMode = modeForm.IsManualSelection;
                    includeLinks = modeForm.IncludeLinkedRooms;
                }

                // --- BƯỚC 2: THU THẬP DỮ LIỆU ---
                List<Element> targetElements = new List<Element>();
                List<RoomInfo> sourceRooms = new List<RoomInfo>(); // Dùng class bao gói để chứa cả Link Room

                ElementId phaseId = doc.ActiveView.get_Parameter(BuiltInParameter.VIEW_PHASE).AsElementId();
                Phase phase = doc.GetElement(phaseId) as Phase;
                if (phase == null)
                {
                    TaskDialog.Show("Error", "Không xác định được Phase của View hiện tại.");
                    return Result.Failed;
                }

                if (isManualMode)
                {
                    try
                    {
                        // Cho phép chọn Room (Host), Link Instance (để lấy Room Link), và Model Elements
                        IList<Reference> mixedRefs = uiDoc.Selection.PickObjects(
                            ObjectType.Element,
                            new MixedSelectionFilter(),
                            "Quét chọn vùng (Room, Link, và Đối tượng cần gán)");

                        foreach (var r in mixedRefs)
                        {
                            Element e = doc.GetElement(r);

                            if (e is Room room && room.Location != null)
                            {
                                sourceRooms.Add(new RoomInfo(room, null));
                            }
                            else if (e is RevitLinkInstance linkInst && includeLinks)
                            {
                                // Nếu chọn Link Instance -> Lấy hết Room trong Link đó
                                var linkRooms = GetRoomsFromLink(linkInst);
                                sourceRooms.AddRange(linkRooms);
                            }
                            else if (e.Category != null && e.Category.CategoryType == CategoryType.Model)
                            {
                                targetElements.Add(e);
                            }
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }
                }
                else // Auto Mode
                {
                    // 1. Room Host
                    var hostRooms = new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .OfClass(typeof(SpatialElement))
                        .Where(e => e is Room)
                        .Cast<Room>()
                        .Where(r => r.Location != null && r.Area > 0);

                    foreach (var r in hostRooms) sourceRooms.Add(new RoomInfo(r, null));

                    // 2. Room Link (Nếu được chọn)
                    if (includeLinks)
                    {
                        var links = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>();
                        foreach (var link in links)
                        {
                            sourceRooms.AddRange(GetRoomsFromLink(link));
                        }
                    }

                    // 3. Target Elements
                    targetElements = new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .WhereElementIsNotElementType()
                        .WhereElementIsViewIndependent()
                        .Where(e => e.Category != null && !(e is Room) && !(e is Space) && e.Category.CategoryType == CategoryType.Model)
                        .ToList();
                }

                if (targetElements.Count == 0)
                {
                    TaskDialog.Show("Info", "Không tìm thấy đối tượng đích.");
                    return Result.Succeeded;
                }
                if (sourceRooms.Count == 0)
                {
                    TaskDialog.Show("Info", "Không tìm thấy Room nào (Host hoặc Link).");
                    return Result.Succeeded;
                }

                // --- BƯỚC 3: CHỌN PARAMETER ---
                List<string> writableParams = GetWritableStringParameters(targetElements.First());
                if (!writableParams.Contains("Comments")) writableParams.Insert(0, "Comments");

                string targetParamName = "";
                using (var paramForm = new ParamSelectionForm(writableParams))
                {
                    if (paramForm.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;
                    targetParamName = paramForm.SelectedParam;
                }

                // --- BƯỚC 4: THỰC THI ---
                int successCount = 0;
                int failCount = 0;
                int skipCount = 0;

                using (Transaction t = new Transaction(doc, "Copy Room Data"))
                {
                    t.Start();

                    // Pre-calc: Để tối ưu tốc độ cho Manual Mode, ta cần biết những Room nào được chọn
                    // Tuy nhiên với Link Room, ID sẽ trùng lặp hoặc không unique trong Host.
                    // Nên ta dùng danh sách sourceRooms để duyệt.

                    foreach (Element elem in targetElements)
                    {
                        Parameter p = elem.LookupParameter(targetParamName);
                        if (p == null || p.IsReadOnly || p.StorageType != StorageType.String)
                        {
                            failCount++; continue;
                        }

                        // Tìm Room chứa Element
                        // Logic mới: Tìm trong danh sách sourceRooms xem Element nằm trong cái nào
                        RoomInfo foundRoomInfo = FindRoomContainingElement(elem, sourceRooms);

                        if (foundRoomInfo != null)
                        {
                            // Nếu tìm thấy trong danh sách sourceRooms -> Có nghĩa là thoả mãn cả 2 điều kiện:
                            // 1. Element nằm trong Room đó.
                            // 2. Room đó nằm trong danh sách đã chọn (Manual) hoặc danh sách toàn bộ (Auto).

                            string roomName = foundRoomInfo.GetRoomName();
                            if (p.Set(roomName)) successCount++;
                            else failCount++;
                        }
                        else
                        {
                            // Không tìm thấy Room nào trong danh sách sourceRooms chứa element này
                            // Có thể element nằm ngoài trời, hoặc nằm trong Room mà user KHÔNG chọn.
                            skipCount++;
                        }
                    }
                    t.Commit();
                }

                string msg = $"Kết quả cập nhật '{targetParamName}':\n" +
                             $"- Thành công: {successCount}\n" +
                             $"- Bỏ qua/Không thuộc Room đã chọn: {skipCount}\n" +
                             $"- Lỗi: {failCount}";
                TaskDialog.Show("Kết quả", msg);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        // --- CORE LOGIC ---

        // Class hỗ trợ xử lý cả Host Room và Link Room
        public class RoomInfo
        {
            public Room RoomObject { get; }
            public RevitLinkInstance LinkInstance { get; }
            public Transform Transform { get; }

            public RoomInfo(Room r, RevitLinkInstance link)
            {
                RoomObject = r;
                LinkInstance = link;
                if (link != null) Transform = link.GetTotalTransform();
                else Transform = Transform.Identity;
            }

            public bool IsPointInRoom(XYZ point)
            {
                // Nếu là Link Room, cần chuyển điểm từ Host Space sang Link Space
                XYZ pointInRoomSpace = Transform.Inverse.OfPoint(point);
                return RoomObject.IsPointInRoom(pointInRoomSpace);
            }

            public string GetRoomName()
            {
                return RoomObject.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
            }
        }

        private List<RoomInfo> GetRoomsFromLink(RevitLinkInstance linkInst)
        {
            var results = new List<RoomInfo>();
            Document linkDoc = linkInst.GetLinkDocument();
            if (linkDoc == null) return results;

            var rooms = new FilteredElementCollector(linkDoc)
                .OfClass(typeof(SpatialElement))
                .Where(e => e is Room)
                .Cast<Room>()
                .Where(r => r.Location != null && r.Area > 0);

            foreach (var r in rooms)
            {
                results.Add(new RoomInfo(r, linkInst));
            }
            return results;
        }

        private RoomInfo FindRoomContainingElement(Element elem, List<RoomInfo> roomsToCheck)
        {
            // Lấy điểm đặt của Element
            XYZ point = null;
            if (elem.Location is LocationPoint lp) point = lp.Point;
            else if (elem is FamilyInstance fi && fi.Location != null && fi.Location is LocationPoint fip) point = fip.Point;
            // Nếu không có Point (vd Wall), lấy tâm BoundingBox
            if (point == null)
            {
                BoundingBoxXYZ bb = elem.get_BoundingBox(null);
                if (bb != null) point = (bb.Min + bb.Max) / 2.0;
            }

            if (point == null) return null;

            // Duyệt qua danh sách Room đã chọn để tìm
            foreach (var rInfo in roomsToCheck)
            {
                if (rInfo.IsPointInRoom(point)) return rInfo;
            }
            return null;
        }

        private List<string> GetWritableStringParameters(Element e)
        {
            var paramsList = new List<string>();
            foreach (Parameter p in e.Parameters)
            {
                if (p.StorageType == StorageType.String && !p.IsReadOnly && p.UserModifiable)
                    paramsList.Add(p.Definition.Name);
            }
            paramsList.Sort();
            return paramsList.Distinct().ToList();
        }

        public class MixedSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem is Room) return true;
                if (elem is RevitLinkInstance) return true; // Cho phép chọn Link
                if (elem.Category != null && elem.Category.CategoryType == CategoryType.Model && !(elem is Space)) return true;
                return false;
            }
            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        // --- UI FORMS ---

        public class ModeSelectionForm : Form
        {
            private RadioButton rbManual;
            private RadioButton rbAll;
            private CheckBox chkIncludeLink; // Checkbox mới
            private Button btnNext;
            private Button btnCancel;

            public bool IsManualSelection => rbManual.Checked;
            public bool IncludeLinkedRooms => chkIncludeLink.Checked;

            public ModeSelectionForm() { SetupUI(); }

            private void SetupUI()
            {
                Text = "SKRevit - Select Mode";
                Size = new Size(420, 320);
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false; MinimizeBox = false;
                BackColor = Color.WhiteSmoke;
                Font = SystemFonts.MessageBoxFont;

                var pnlHeader = CreateHeaderPanel();
                pnlHeader.Dock = DockStyle.Top; pnlHeader.Height = 45;

                var pnlContent = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
                var gb = new GroupBox { Text = "Selection Mode", Dock = DockStyle.Top, Height = 140 };

                rbManual = new RadioButton { Text = "Option 1: Quét chọn vùng (Room/Link + Đối tượng)", Location = new Point(20, 30), AutoSize = true, Checked = true };
                rbAll = new RadioButton { Text = "Option 2: Tự động toàn bộ (Active View)", Location = new Point(20, 60), AutoSize = true };

                chkIncludeLink = new CheckBox
                {
                    Text = "Include Linked Rooms (Bao gồm Room trong Link)",
                    Location = new Point(40, 95),
                    AutoSize = true,
                    Checked = true,
                    ForeColor = Color.DarkBlue
                };

                gb.Controls.Add(rbManual); gb.Controls.Add(rbAll); gb.Controls.Add(chkIncludeLink);
                pnlContent.Controls.Add(gb);

                var pnlBtn = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 50, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(10) };
                btnNext = new Button { Text = "Next >", Width = 90, Height = 30, BackColor = Color.White, DialogResult = DialogResult.OK };
                btnCancel = new Button { Text = "Cancel", Width = 90, Height = 30, BackColor = Color.White, DialogResult = DialogResult.Cancel };
                pnlBtn.Controls.Add(btnNext); pnlBtn.Controls.Add(btnCancel);

                Controls.Add(pnlContent); Controls.Add(pnlBtn); Controls.Add(pnlHeader);
                AcceptButton = btnNext; CancelButton = btnCancel;
            }
        }

        public class ParamSelectionForm : Form
        {
            private ComboBox cbParam;
            private Button btnRun, btnCancel;
            public string SelectedParam => cbParam.Text;

            public ParamSelectionForm(List<string> paramsList) { SetupUI(paramsList); }

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
                pnlHeader.Dock = DockStyle.Top; pnlHeader.Height = 45;

                var pnlContent = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
                var lbl = new Label { Text = "Chọn tham số cần điền tên Room vào:", Dock = DockStyle.Top, Height = 25 };
                cbParam = new ComboBox { DataSource = paramsList, Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDown, AutoCompleteMode = AutoCompleteMode.SuggestAppend, AutoCompleteSource = AutoCompleteSource.ListItems };
                if (paramsList.Contains("Comments")) cbParam.SelectedItem = "Comments";
                pnlContent.Controls.Add(cbParam); pnlContent.Controls.Add(lbl);

                var pnlBtn = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 50, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(10) };
                btnRun = new Button { Text = "Run", Width = 90, Height = 30, BackColor = Color.White, DialogResult = DialogResult.OK };
                btnCancel = new Button { Text = "Cancel", Width = 90, Height = 30, BackColor = Color.White, DialogResult = DialogResult.Cancel };
                pnlBtn.Controls.Add(btnRun); pnlBtn.Controls.Add(btnCancel);

                Controls.Add(pnlContent); Controls.Add(pnlBtn); Controls.Add(pnlHeader);
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