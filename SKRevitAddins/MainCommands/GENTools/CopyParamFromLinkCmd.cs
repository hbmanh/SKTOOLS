using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using SKRevitAddins.Utils;
using Application = Autodesk.Revit.ApplicationServices.Application;
using ComboBox = System.Windows.Forms.ComboBox;
using Form = System.Windows.Forms.Form;
using TextBox = System.Windows.Forms.TextBox;
using View = Autodesk.Revit.DB.View;
using Point = System.Drawing.Point;
using Color = System.Drawing.Color;
using Panel = System.Windows.Forms.Panel;

namespace SKRevitAddins.GENTools
{
    [Transaction(TransactionMode.Manual)]
    public class CopyParamFromLinkCmd : IExternalCommand
    {
        public class ParamData
        {
            public string Value { get; set; }
            public bool IsReadOnly { get; set; }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                Reference pickedRef = null;
                try
                {
                    pickedRef = uiDoc.Selection.PickObject(ObjectType.LinkedElement, "Pick Linked Element (Source)");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                RevitLinkInstance linkInstance = doc.GetElement(pickedRef.ElementId) as RevitLinkInstance;
                Document linkDoc = linkInstance.GetLinkDocument();
                Element linkedElement = linkDoc.GetElement(pickedRef.LinkedElementId);

                var sourceParamData = GetElementParametersData(linkedElement);

                IList<Reference> targetRefs = null;
                try
                {
                    targetRefs = uiDoc.Selection.PickObjects(ObjectType.Element, "Select Target Elements (Destination)");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (targetRefs.Count == 0) return Result.Cancelled;

                Element firstTarget = doc.GetElement(targetRefs.First());
                List<string> availableTargetParams = GetWritableParamNames(firstTarget);

                string sourceParamName = "";
                string targetParamName = "";
                bool copyAsText = false;

                using (var form = new ParamMappingForm(sourceParamData, availableTargetParams, linkedElement.Name, linkInstance.Name, targetRefs.Count))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;

                    sourceParamName = form.SelectedSourceParam;
                    targetParamName = form.TargetParamName;
                    copyAsText = form.IsCopyAsText;
                }

                Parameter srcParam = linkedElement.LookupParameter(sourceParamName);
                if (srcParam == null) return Result.Failed;

                int successCount = 0;
                int failCount = 0;

                using (Transaction t = new Transaction(doc, "Copy Linked Param"))
                {
                    t.Start();
                    foreach (var refT in targetRefs)
                    {
                        Element targetEl = doc.GetElement(refT);
                        Parameter targetParam = targetEl.LookupParameter(targetParamName);

                        if (targetParam == null || targetParam.IsReadOnly)
                        {
                            failCount++;
                            continue;
                        }

                        if (SetParameterValue(srcParam, targetParam, copyAsText))
                            successCount++;
                        else
                            failCount++;
                    }
                    t.Commit();
                }

                TaskDialog.Show("Result", $"Copy '{sourceParamName}' -> '{targetParamName}'\nSource Value: {GetParamValueString(srcParam)}\n\nSuccess: {successCount}\nFailed/ReadOnly: {failCount}");
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        private List<string> GetWritableParamNames(Element e)
        {
            var list = new List<string>();
            foreach (Parameter p in e.Parameters)
            {
                if (!p.IsReadOnly)
                {
                    list.Add(p.Definition.Name);
                }
            }
            list.Sort();
            return list;
        }

        private Dictionary<string, ParamData> GetElementParametersData(Element e)
        {
            var data = new Dictionary<string, ParamData>();
            foreach (Parameter p in e.Parameters)
            {
                string pName = p.Definition.Name;
                string pValue = GetParamValueString(p);
                bool isRO = p.IsReadOnly;

                if (!data.ContainsKey(pName))
                {
                    data.Add(pName, new ParamData { Value = pValue, IsReadOnly = isRO });
                }
            }
            return data.OrderBy(k => k.Key).ToDictionary(k => k.Key, k => k.Value);
        }

        private string GetParamValueString(Parameter p)
        {
            if (p == null) return "";
            string val = p.AsValueString();
            if (string.IsNullOrEmpty(val)) { try { val = p.AsString(); } catch { } }
            return val ?? "";
        }

        private bool SetParameterValue(Parameter source, Parameter target, bool forceText)
        {
            try
            {
                if (forceText || target.StorageType == StorageType.String)
                    return target.Set(GetParamValueString(source));

                if (source.StorageType == target.StorageType)
                {
                    switch (source.StorageType)
                    {
                        case StorageType.Double: return target.Set(source.AsDouble());
                        case StorageType.Integer: return target.Set(source.AsInteger());
                        case StorageType.ElementId: return target.Set(source.AsValueString());
                        default: return target.Set(GetParamValueString(source));
                    }
                }
                else
                {
                    string sVal = GetParamValueString(source);
                    if (target.StorageType == StorageType.Double && double.TryParse(sVal, out double d)) return target.Set(d);
                    if (target.StorageType == StorageType.Integer && int.TryParse(sVal, out int i)) return target.Set(i);
                }
            }
            catch { return false; }
            return false;
        }

        public class ParamMappingForm : Form
        {
            private ComboBox cbSourceParam;
            private TextBox txtSourceValue;
            private ComboBox cbTargetParam;
            private CheckBox chkForceText;
            private Button btnOK;
            private Button btnCancel;

            private Dictionary<string, ParamData> _sourceData;

            public string SelectedSourceParam => cbSourceParam.Text;
            public string TargetParamName => cbTargetParam.Text;
            public bool IsCopyAsText => chkForceText.Checked;

            public ParamMappingForm(Dictionary<string, ParamData> sourceData, List<string> targetParams, string elementName, string linkName, int targetCount)
            {
                _sourceData = sourceData;

                Text = "SKRevit - Copy Parameter Data";
                MinimumSize = new Size(800, 450);
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                BackColor = Color.WhiteSmoke;
                Font = SystemFonts.MessageBoxFont;

                var mainLayout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = 5,
                    Padding = new Padding(0)
                };

                mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));
                mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var headerPanel = new Panel
                {
                    Dock = DockStyle.Fill,
                    BackColor = Color.White,
                    Margin = new Padding(0)
                };

                var logoBox = new PictureBox
                {
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Size = new Size(30, 30),
                    Location = new Point(10, 7)
                };
                try { LogoHelper.TryLoadLogo(logoBox); } catch { }
                headerPanel.Controls.Add(logoBox);

                var companyLabel = new Label
                {
                    Text = "Shinken Group®",
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                    AutoSize = true,
                    Location = new Point(50, 12),
                    ForeColor = Color.Black
                };
                headerPanel.Controls.Add(companyLabel);

                var lblInfo = new Label
                {
                    Text = $"Source: {elementName} ({linkName})\nTargets: {targetCount} elements selected",
                    AutoSize = true,
                    Font = new Font(Font, FontStyle.Bold),
                    ForeColor = Color.DimGray,
                    Margin = new Padding(15, 15, 15, 10)
                };

                var inputPanel = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 2,
                    RowCount = 4,
                    AutoSize = true,
                    Padding = new Padding(15, 0, 15, 0)
                };
                inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F));
                inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                inputPanel.Controls.Add(new Label { Text = "Source Param:", Anchor = AnchorStyles.Left | AnchorStyles.Top, AutoSize = true, Padding = new Padding(0, 6, 0, 0) }, 0, 0);
                cbSourceParam = new ComboBox
                {
                    DataSource = new List<string>(_sourceData.Keys),
                    Anchor = AnchorStyles.Left | AnchorStyles.Right,
                    DropDownStyle = ComboBoxStyle.DropDown,
                    AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                    AutoCompleteSource = AutoCompleteSource.ListItems
                };
                inputPanel.Controls.Add(cbSourceParam, 1, 0);

                inputPanel.Controls.Add(new Label { Text = "Review Value:", Anchor = AnchorStyles.Left | AnchorStyles.Top, AutoSize = true, Padding = new Padding(0, 6, 0, 0), ForeColor = Color.Blue }, 0, 1);
                txtSourceValue = new TextBox
                {
                    ReadOnly = true,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right,
                    BackColor = Color.WhiteSmoke,
                    ForeColor = Color.Blue
                };
                inputPanel.Controls.Add(txtSourceValue, 1, 1);

                inputPanel.Controls.Add(new Label { Text = "Target Param:", Anchor = AnchorStyles.Left | AnchorStyles.Top, AutoSize = true, Padding = new Padding(0, 6, 0, 0) }, 0, 2);
                cbTargetParam = new ComboBox
                {
                    DataSource = targetParams,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right,
                    DropDownStyle = ComboBoxStyle.DropDown,
                    AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                    AutoCompleteSource = AutoCompleteSource.ListItems
                };

                inputPanel.Controls.Add(cbTargetParam, 1, 2);

                var btnSyncName = new Button { Text = "⇩ Copy Name Down", AutoSize = true, Margin = new Padding(0, 5, 0, 0), BackColor = Color.White };
                btnSyncName.Click += (s, e) => { cbTargetParam.Text = cbSourceParam.Text; };
                inputPanel.Controls.Add(btnSyncName, 1, 3);

                chkForceText = new CheckBox
                {
                    Text = "Force Convert to Text",
                    AutoSize = true,
                    Checked = true,
                    Margin = new Padding(15, 10, 15, 10)
                };

                var btnPanel = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.RightToLeft,
                    Dock = DockStyle.Fill,
                    AutoSize = true,
                    Padding = new Padding(15)
                };
                btnOK = new Button { Text = "Run", AutoSize = true, DialogResult = DialogResult.OK, Height = 35, BackColor = Color.White, Width = 80 };
                btnCancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel, Height = 35, BackColor = Color.White, Width = 80 };
                btnPanel.Controls.Add(btnOK);
                btnPanel.Controls.Add(btnCancel);

                mainLayout.Controls.Add(headerPanel, 0, 0);
                mainLayout.Controls.Add(lblInfo, 0, 1);
                mainLayout.Controls.Add(inputPanel, 0, 2);
                mainLayout.Controls.Add(chkForceText, 0, 3);
                mainLayout.Controls.Add(btnPanel, 0, 4);

                Controls.Add(mainLayout);
                AcceptButton = btnOK;
                CancelButton = btnCancel;

                cbSourceParam.TextChanged += (s, e) => {
                    UpdateValueReview();
                };
                cbSourceParam.SelectedIndexChanged += (s, e) => {
                    UpdateValueReview();
                    if (string.IsNullOrWhiteSpace(cbTargetParam.Text))
                        cbTargetParam.Text = cbSourceParam.Text;
                };

                if (cbSourceParam.Items.Count > 0) cbSourceParam.SelectedIndex = 0;
            }

            private void UpdateValueReview()
            {
                string key = cbSourceParam.Text;
                if (_sourceData.ContainsKey(key))
                    txtSourceValue.Text = _sourceData[key].Value;
                else
                    txtSourceValue.Text = "";
            }
        }
    }
}