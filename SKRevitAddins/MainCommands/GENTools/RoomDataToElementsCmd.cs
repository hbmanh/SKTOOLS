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
                bool isManualMode = true;
                using (var modeForm = new ModeSelectionForm())
                {
                    if (modeForm.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;
                    isManualMode = modeForm.IsManualSelection;
                }

                List<Element> targetElements = new List<Element>();
                List<Room> sourceRooms = new List<Room>();

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
                        IList<Reference> mixedRefs = uiDoc.Selection.PickObjects(
                            ObjectType.Element,
                            new MixedSelectionFilter(),
                            "Quét chọn vùng chứa cả Room và Đối tượng cần gán");

                        foreach (var r in mixedRefs)
                        {
                            Element e = doc.GetElement(r);

                            if (e is Room room && room.Location != null)
                            {
                                sourceRooms.Add(room);
                            }
                            else if (e.Category != null && e.Category.CategoryType == CategoryType.Model)
                            {
                                targetElements.Add(e);
                            }
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }
                }
                else 
                {
                    sourceRooms = new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .OfClass(typeof(SpatialElement))
                        .Where(e => e is Room)
                        .Cast<Room>()
                        .Where(r => r.Location != null && r.Area > 0)
                        .ToList();

                    targetElements = new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .WhereElementIsNotElementType()
                        .WhereElementIsViewIndependent()
                        .Where(e => e.Category != null && !(e is Room) && !(e is Space) && e.Category.CategoryType == CategoryType.Model)
                        .ToList();
                }

                if (targetElements.Count == 0)
                {
                    TaskDialog.Show("Info", "Không tìm thấy đối tượng đích (Thiết bị/Cửa/Cột...) nào trong vùng chọn.");
                    return Result.Succeeded;
                }
                if (sourceRooms.Count == 0)
                {
                    TaskDialog.Show("Info", "Không tìm thấy Room nào trong vùng chọn.");
                    return Result.Succeeded;
                }

                List<string> writableParams = GetWritableStringParameters(targetElements.First());
                if (!writableParams.Contains("Comments")) writableParams.Insert(0, "Comments");
                if (!writableParams.Contains("Mark")) writableParams.Insert(1, "Mark");

                string targetParamName = "";
                using (var paramForm = new ParamSelectionForm(writableParams))
                {
                    if (paramForm.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;
                    targetParamName = paramForm.SelectedParam;
                }

                int successCount = 0;
                int failCount = 0;
                int skipCount = 0;

                using (Transaction t = new Transaction(doc, "Copy Room Name to Elements"))
                {
                    t.Start();

                    HashSet<ElementId> selectedRoomIds = new HashSet<ElementId>(sourceRooms.Select(r => r.Id));

                    foreach (Element elem in targetElements)
                    {
                        Parameter p = elem.LookupParameter(targetParamName);
                        if (p == null || p.IsReadOnly || p.StorageType != StorageType.String)
                        {
                            failCount++;
                            continue;
                        }

                        Room foundRoom = GetRoomOfElement(elem, phase);

                        if (foundRoom != null)
                        {
                            if (isManualMode && !selectedRoomIds.Contains(foundRoom.Id))
                            {
                                skipCount++;
                                continue;
                            }

                            string roomName = foundRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();

                            if (p.Set(roomName)) successCount++;
                            else failCount++;
                        }
                        else
                        {
                            // Đối tượng nằm ngoài Room
                            failCount++;
                        }
                    }
                    t.Commit();
                }

                string msg = $"Hoàn tất!\n\n" +
                             $"- Đã gán '{targetParamName}': {successCount}\n" +
                             $"- Bỏ qua (Khác Room đã chọn): {skipCount}\n" +
                             $"- Lỗi (Không Param/Không Room): {failCount}";

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

        public class MixedSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem is Room) return true;

                if (elem.Category != null && elem.Category.CategoryType == CategoryType.Model && !(elem is Space))
                    return true;

                return false;
            }
            public bool AllowReference(Reference reference, XYZ position) => false;
        }


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

                // Cập nhật text hướng dẫn cho đúng với logic mới
                rbManual = new RadioButton { Text = "Option 1: Quét chọn vùng (Room + Đối tượng)", Location = new Point(20, 30), AutoSize = true, Checked = true };
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
                Size = new Size(450, 220);
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
                    DropDownStyle = ComboBoxStyle.DropDown,
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