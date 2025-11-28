using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SKRevitAddins.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using AnchorStyles = System.Windows.Forms.AnchorStyles;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Button = System.Windows.Forms.Button;
using Color = System.Drawing.Color;
using ContentAlignment = System.Drawing.ContentAlignment;
using DataGridView = System.Windows.Forms.DataGridView;
using DataGridViewAutoSizeColumnMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode;
using DataGridViewButtonColumn = System.Windows.Forms.DataGridViewButtonColumn;
using DataGridViewCheckBoxColumn = System.Windows.Forms.DataGridViewCheckBoxColumn;
using DataGridViewDataErrorContexts = System.Windows.Forms.DataGridViewDataErrorContexts;
using DataGridViewRow = System.Windows.Forms.DataGridViewRow;
using DataGridViewSelectionMode = System.Windows.Forms.DataGridViewSelectionMode;
using DataGridViewTextBoxColumn = System.Windows.Forms.DataGridViewTextBoxColumn;
using DockStyle = System.Windows.Forms.DockStyle;
using Font = System.Drawing.Font;
using FontStyle = System.Drawing.FontStyle;
using Form = System.Windows.Forms.Form;
using IWin32Window = System.Windows.Forms.IWin32Window;
using Label = System.Windows.Forms.Label;
using Padding = System.Windows.Forms.Padding;
using Panel = System.Windows.Forms.Panel;
using PictureBox = System.Windows.Forms.PictureBox;
using PictureBoxSizeMode = System.Windows.Forms.PictureBoxSizeMode;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

namespace SKRevitAddins.GENTools
{
    [Transaction(TransactionMode.Manual)]
    public class CheckDuplicateElementsCmd : IExternalCommand
    {
        public static DuplicateCheckForm _instanceForm;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (_instanceForm != null && !_instanceForm.IsDisposed)
                {
                    _instanceForm.BringToFront();
                    return Result.Succeeded;
                }

                DuplicateRequestHandler handler = new DuplicateRequestHandler();
                ExternalEvent exEvent = ExternalEvent.Create(handler);

                _instanceForm = new DuplicateCheckForm(exEvent, handler, commandData.Application.ActiveUIDocument);
                handler.SetForm(_instanceForm);

                _instanceForm.FormClosed += (s, e) => { _instanceForm = null; };

                IntPtr revitWindowHandle = Process.GetCurrentProcess().MainWindowHandle;
                WindowWrapper wrapper = new WindowWrapper(revitWindowHandle);

                _instanceForm.Show(wrapper);

                handler.MakeRequest(RequestId.Reload);
                exEvent.Raise();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.ToString());
                return Result.Failed;
            }
        }

        public class WindowWrapper : IWin32Window
        {
            private IntPtr _handle;
            public WindowWrapper(IntPtr handle) { _handle = handle; }
            public IntPtr Handle => _handle;
        }
    }

    public class DuplicateGroup
    {
        public Element KeeperElement { get; set; }
        public List<Element> DuplicateElements { get; set; }
    }

    public enum RequestId
    {
        None,
        Reload,
        Delete
    }

    public class DuplicateRequestHandler : IExternalEventHandler
    {
        private RequestId _request = RequestId.None;
        private DuplicateCheckForm _form;
        private bool _deleteSelectedOnly = false;

        public void SetForm(DuplicateCheckForm form) { _form = form; }

        public void MakeRequest(RequestId request, bool deleteSelectedOnly = false)
        {
            _request = request;
            _deleteSelectedOnly = deleteSelectedOnly;
        }

        public string GetName() => "Duplicate Check Handler";

        public void Execute(UIApplication app)
        {
            try
            {
                UIDocument uiDoc = app.ActiveUIDocument;
                Document doc = uiDoc.Document;

                switch (_request)
                {
                    case RequestId.Reload:
                        RunCheckLogic(uiDoc);
                        break;
                    case RequestId.Delete:
                        DeleteLogic(doc);
                        RunCheckLogic(uiDoc);
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Handler Error", ex.Message);
            }
        }

        private void RunCheckLogic(UIDocument uiDoc)
        {
            if (_form == null || _form.IsDisposed) return;

            Document doc = uiDoc.Document;
            List<Element> elementsToCheck = new List<Element>();
            ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();

            if (selectedIds.Count > 0)
            {
                foreach (ElementId id in selectedIds) elementsToCheck.Add(doc.GetElement(id));
            }
            else
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                elementsToCheck = collector.WhereElementIsNotElementType()
                                           .Where(e => e.Category != null && e.Location != null)
                                           .ToList();
            }

            double tolerance = 100.0 / 304.8;
            var duplicates = FindDuplicates(elementsToCheck, tolerance);
            _form.UpdateData(duplicates);
        }

        private void DeleteLogic(Document doc)
        {
            List<ElementId> idsToDelete = _form.GetIdsToDelete(_deleteSelectedOnly);
            if (idsToDelete.Count > 0)
            {
                using (Transaction t = new Transaction(doc, "Delete Duplicates"))
                {
                    t.Start();
                    doc.Delete(idsToDelete);
                    t.Commit();
                }
            }
        }

        private List<DuplicateGroup> FindDuplicates(List<Element> elements, double tolerance)
        {
            var results = new List<DuplicateGroup>();
            var processedIds = new HashSet<ElementId>();
            var groupedByType = elements.GroupBy(e => e.GetTypeId()).ToList();

            foreach (var group in groupedByType)
            {
                var elemsInType = group.ToList();
                for (int i = 0; i < elemsInType.Count; i++)
                {
                    Element e1 = elemsInType[i];
                    if (processedIds.Contains(e1.Id)) continue;
                    XYZ p1 = GetElementLocation(e1);
                    if (p1 == null) continue;

                    var currentGroup = new DuplicateGroup { KeeperElement = e1, DuplicateElements = new List<Element>() };

                    for (int j = i + 1; j < elemsInType.Count; j++)
                    {
                        Element e2 = elemsInType[j];
                        if (processedIds.Contains(e2.Id)) continue;
                        XYZ p2 = GetElementLocation(e2);
                        if (p2 == null) continue;

                        if (p1.DistanceTo(p2) <= tolerance)
                        {
                            currentGroup.DuplicateElements.Add(e2);
                            processedIds.Add(e2.Id);
                        }
                    }

                    if (currentGroup.DuplicateElements.Count > 0)
                    {
                        results.Add(currentGroup);
                        processedIds.Add(e1.Id);
                    }
                }
            }
            return results;
        }

        private XYZ GetElementLocation(Element e)
        {
            if (e.Location is LocationPoint lp) return lp.Point;
            if (e.Location is LocationCurve lc) return (lc.Curve.GetEndPoint(0) + lc.Curve.GetEndPoint(1)) / 2.0;
            BoundingBoxXYZ bbox = e.get_BoundingBox(null);
            return bbox != null ? (bbox.Min + bbox.Max) / 2.0 : null;
        }
    }

    public class DuplicateCheckForm : Form
    {
        private DataGridView dgvDuplicates;
        private Button btnDeleteAll, btnDeleteSelected, btnCancel, btnReload;
        private Label lblInfo;
        private ExternalEvent _exEvent;
        private DuplicateRequestHandler _handler;
        private UIDocument _uiDoc;

        public DuplicateCheckForm(ExternalEvent exEvent, DuplicateRequestHandler handler, UIDocument uiDoc)
        {
            _exEvent = exEvent;
            _handler = handler;
            _uiDoc = uiDoc;
            InitializeComponent();
        }

        public void UpdateData(List<DuplicateGroup> groups)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<List<DuplicateGroup>>(UpdateData), groups);
                return;
            }

            dgvDuplicates.Rows.Clear();
            lblInfo.Text = $"Found {groups.Count} locations with duplicates.";

            foreach (var group in groups)
            {
                int index = dgvDuplicates.Rows.Add();
                var row = dgvDuplicates.Rows[index];

                row.Cells["colCheck"].Value = false;
                row.Cells["colID"].Value = group.KeeperElement.Id.ToString();

                string displayName = group.KeeperElement.Name;
                ElementId typeId = group.KeeperElement.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    ElementType type = group.KeeperElement.Document.GetElement(typeId) as ElementType;
                    if (type != null)
                    {
                        displayName = $"{type.FamilyName} : {type.Name}";
                    }
                }

                row.Cells["colName"].Value = displayName;
                row.Cells["colCount"].Value = group.DuplicateElements.Count;
                row.Tag = group;
            }
        }

        public List<ElementId> GetIdsToDelete(bool selectedOnly)
        {
            dgvDuplicates.EndEdit();

            List<ElementId> ids = new List<ElementId>();
            foreach (DataGridViewRow row in dgvDuplicates.Rows)
            {
                if (row.Tag is DuplicateGroup group)
                {
                    if (selectedOnly)
                    {
                        object val = row.Cells["colCheck"].Value;
                        bool isChecked = (val != null && (bool)val == true);
                        if (isChecked)
                        {
                            ids.AddRange(group.DuplicateElements.Select(e => e.Id));
                        }
                    }
                    else
                    {
                        ids.AddRange(group.DuplicateElements.Select(e => e.Id));
                    }
                }
            }
            return ids;
        }

        private void InitializeComponent()
        {
            Text = "SKRevit - Duplicate Check (Modeless)";
            Size = new Size(800, 500);
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            TopMost = true;
            BackColor = Color.WhiteSmoke;
            Font = System.Drawing.SystemFonts.MessageBoxFont;

            var headerPanel = CreateHeaderPanel();
            headerPanel.Dock = DockStyle.Top;
            headerPanel.Height = 45;

            var actionPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 45,
                Padding = new Padding(10, 5, 10, 5),
                BackColor = Color.WhiteSmoke
            };

            btnReload = new Button
            {
                Text = "Reload",
                Width = 100,
                Dock = DockStyle.Right,
                BackColor = Color.LightBlue,
                Cursor = System.Windows.Forms.Cursors.Hand
            };
            btnReload.Click += (s, e) => {
                lblInfo.Text = "Checking...";
                _handler.MakeRequest(RequestId.Reload);
                _exEvent.Raise();
            };

            lblInfo = new Label
            {
                Text = "Initializing...",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            actionPanel.Controls.Add(btnReload);
            actionPanel.Controls.Add(lblInfo);

            var botPanel = new Panel { Dock = DockStyle.Bottom, Height = 50, Padding = new Padding(10) };

            btnDeleteAll = new Button
            {
                Text = "Delete ALL",
                Width = 200,
                BackColor = Color.IndianRed,
                ForeColor = Color.White,
                Dock = DockStyle.Left
            };
            btnDeleteAll.Click += (s, e) => {
                _handler.MakeRequest(RequestId.Delete, false);
                _exEvent.Raise();
            };

            btnDeleteSelected = new Button
            {
                Text = "Delete Checked Only",
                Width = 150,
                Left = 210,
                Dock = DockStyle.Left
            };
            btnDeleteSelected.Click += (s, e) => {
                _handler.MakeRequest(RequestId.Delete, true);
                _exEvent.Raise();
            };

            btnCancel = new Button { Text = "Close", Width = 100, Dock = DockStyle.Right };
            btnCancel.Click += (s, e) => Close();

            Panel spacer = new Panel { Width = 10, Dock = DockStyle.Left };
            botPanel.Controls.Add(btnDeleteSelected);
            botPanel.Controls.Add(spacer);
            botPanel.Controls.Add(btnDeleteAll);
            botPanel.Controls.Add(btnCancel);

            dgvDuplicates = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false, AllowUserToAddRows = false, RowHeadersVisible = false, 
                SelectionMode = DataGridViewSelectionMode.FullRowSelect, BackgroundColor = Color.White,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing,
                ColumnHeadersHeight = 45,
            };

            dgvDuplicates.CurrentCellDirtyStateChanged += (s, e) => {
                if (dgvDuplicates.IsCurrentCellDirty)
                    dgvDuplicates.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            dgvDuplicates.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Del?", Width = 60, Name = "colCheck" });
            dgvDuplicates.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Keeper ID", Width = 100, Name = "colID", ReadOnly = true });
            dgvDuplicates.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Family : Type", Width = 200, Name = "colName", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            dgvDuplicates.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Count", Width = 60, Name = "colCount", ReadOnly = true });

            var colBtn = new DataGridViewButtonColumn { HeaderText = "Review", Text = "Select", UseColumnTextForButtonValue = true, Width = 80, Name = "colBtn" };
            dgvDuplicates.Columns.Add(colBtn);

            dgvDuplicates.CellContentClick += (s, e) => {
                if (e.RowIndex >= 0 && e.ColumnIndex == dgvDuplicates.Columns["colBtn"].Index)
                {
                    if (dgvDuplicates.Rows[e.RowIndex].Tag is DuplicateGroup group)
                    {
                        var ids = new List<ElementId> { group.KeeperElement.Id };
                        ids.AddRange(group.DuplicateElements.Select(x => x.Id));
                        _uiDoc.Selection.SetElementIds(ids);
                        _uiDoc.ShowElements(ids);
                    }
                }
            };

            Controls.Add(dgvDuplicates);
            Controls.Add(botPanel);
            Controls.Add(actionPanel);
            Controls.Add(headerPanel);
        }

        private Panel CreateHeaderPanel()
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