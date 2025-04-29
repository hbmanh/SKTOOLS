using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Structure;
using ComboBox = System.Windows.Forms.ComboBox;
using Form = System.Windows.Forms.Form;
using View = Autodesk.Revit.DB.View;

namespace CopyViewTemplateFromLink
{
    [Transaction(TransactionMode.Manual)]
    public class CopyViewTemplateCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Lấy các Link Document
            List<Document> linkedDocs = GetLinkedDocuments(doc);

            if (linkedDocs.Count == 0)
            {
                TaskDialog.Show("Thông báo", "Không có file link nào.");
                return Result.Cancelled;
            }

            // Mở form chọn file Link và ViewTemplate
            using (var form = new LinkAndViewTemplateSelectorForm(linkedDocs))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    Document linkedDoc = form.SelectedLinkDocument;
                    var selectedTemplates = form.SelectedTemplates;

                    if (linkedDoc == null || selectedTemplates.Count == 0)
                    {
                        TaskDialog.Show("Thông báo", "Chưa chọn file link hoặc ViewTemplate.");
                        return Result.Cancelled;
                    }

                    using (Transaction trans = new Transaction(doc, "Copy ViewTemplates"))
                    {
                        trans.Start();

                        foreach (var vt in selectedTemplates)
                        {
                            CopyElement(doc, linkedDoc, vt.Id);
                        }

                        trans.Commit();
                    }

                    TaskDialog.Show("Thành công", $"Đã copy {selectedTemplates.Count} ViewTemplate(s).");
                }
            }

            return Result.Succeeded;
        }

        private List<Document> GetLinkedDocuments(Document doc)
        {
            List<Document> linkedDocs = new List<Document>();
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>();

            foreach (var linkInstance in linkInstances)
            {
                Document linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc != null)
                    linkedDocs.Add(linkedDoc);
            }

            return linkedDocs;
        }

        private void CopyElement(Document targetDoc, Document sourceDoc, ElementId elementId)
        {
            ICollection<ElementId> elementIds = new List<ElementId>() { elementId };
            ElementTransformUtils.CopyElements(sourceDoc, elementIds, targetDoc, null, new CopyPasteOptions());
        }
    }

    // ---------------- Form chọn Link Document và View Template --------------------
    public class LinkAndViewTemplateSelectorForm : Form
    {
        private ComboBox comboBoxLinks;
        private CheckedListBox checkedListBoxTemplates;
        private Button loadButton;
        private Button okButton;
        private Button cancelButton;

        private List<Document> linkDocs;
        private Dictionary<string, Document> linkNameToDocMap;

        public Document SelectedLinkDocument { get; private set; }
        public List<View> SelectedTemplates { get; private set; } = new List<View>();

        public LinkAndViewTemplateSelectorForm(List<Document> linkDocuments)
        {
            this.linkDocs = linkDocuments;
            this.linkNameToDocMap = new Dictionary<string, Document>();

            foreach (var doc in linkDocuments)
            {
                string docName = doc.Title;
                linkNameToDocMap[docName] = doc;
            }

            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.comboBoxLinks = new ComboBox();
            this.checkedListBoxTemplates = new CheckedListBox();
            this.loadButton = new Button();
            this.okButton = new Button();
            this.cancelButton = new Button();

            this.SuspendLayout();

            // comboBoxLinks
            this.comboBoxLinks.DropDownStyle = ComboBoxStyle.DropDownList;
            this.comboBoxLinks.Dock = DockStyle.Top;
            this.comboBoxLinks.Height = 30;
            this.comboBoxLinks.Items.AddRange(linkNameToDocMap.Keys.ToArray());

            // loadButton
            this.loadButton.Text = "Load ViewTemplates";
            this.loadButton.Dock = DockStyle.Top;
            this.loadButton.Height = 30;
            this.loadButton.Click += new EventHandler(this.LoadButton_Click);

            // checkedListBoxTemplates
            this.checkedListBoxTemplates.FormattingEnabled = true;
            this.checkedListBoxTemplates.Dock = DockStyle.Top;
            this.checkedListBoxTemplates.Height = 300;

            // okButton
            this.okButton.Text = "OK";
            this.okButton.Dock = DockStyle.Bottom;
            this.okButton.Click += new EventHandler(this.OkButton_Click);

            // cancelButton
            this.cancelButton.Text = "Cancel";
            this.cancelButton.Dock = DockStyle.Bottom;
            this.cancelButton.Click += new EventHandler(this.CancelButton_Click);

            // Form
            this.Controls.Add(this.checkedListBoxTemplates);
            this.Controls.Add(this.loadButton);
            this.Controls.Add(this.comboBoxLinks);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.cancelButton);
            this.Text = "Chọn File Link và View Templates";
            this.Height = 500;
            this.Width = 400;
            this.StartPosition = FormStartPosition.CenterScreen;

            this.ResumeLayout(false);
        }

        private void LoadButton_Click(object sender, EventArgs e)
        {
            checkedListBoxTemplates.Items.Clear();

            if (comboBoxLinks.SelectedItem == null)
            {
                MessageBox.Show("Hãy chọn file link trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string selectedLinkName = comboBoxLinks.SelectedItem.ToString();

            if (!linkNameToDocMap.TryGetValue(selectedLinkName, out Document selectedDoc))
            {
                MessageBox.Show("Không tìm thấy file link.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            this.SelectedLinkDocument = selectedDoc;

            var templates = new FilteredElementCollector(selectedDoc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.IsTemplate)
                .ToList();

            foreach (var template in templates)
            {
                checkedListBoxTemplates.Items.Add(new ViewTemplateWrapper(template), false);
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            if (SelectedLinkDocument == null)
            {
                MessageBox.Show("Bạn chưa chọn file link và load view template.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (var item in checkedListBoxTemplates.CheckedItems)
            {
                if (item is ViewTemplateWrapper wrapper)
                {
                    SelectedTemplates.Add(wrapper.View);
                }
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private class ViewTemplateWrapper
        {
            public View View { get; private set; }

            public ViewTemplateWrapper(View view)
            {
                this.View = view;
            }

            public override string ToString()
            {
                return View.Name;
            }
        }
    }
}
