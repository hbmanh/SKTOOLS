#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Form = System.Windows.Forms.Form;
#endregion

namespace SKRevitAddins.GENTools
{
    [Transaction(TransactionMode.Manual)]
    public class DisallowBeamJoinsCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Lấy tất cả dầm (structural framing)
                var beams = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .OfType<FamilyInstance>()
                    .ToList();

                if (beams.Count == 0)
                {
                    TaskDialog.Show("Info", "Không tìm thấy dầm (Structural Framing) nào trong model.");
                    return Result.Succeeded;
                }

                // Hiển thị form progress đơn giản
                using (ProgressForm pf = new ProgressForm(beams.Count))
                {
                    pf.Show();
                    pf.Refresh();

                    using (Transaction t = new Transaction(doc, "Disallow Joins on All Beams"))
                    {
                        t.Start();

                        int processed = 0;
                        foreach (var beam in beams)
                        {
                            try
                            {
                                // 0 = start (begin), 1 = end
                                // StructuralFramingUtils.DisallowJoinAtEnd sẽ
                                // ngắt join nếu đang được joined và đặt end là disallowed.
                                StructuralFramingUtils.DisallowJoinAtEnd(beam, 0);
                                StructuralFramingUtils.DisallowJoinAtEnd(beam, 1);
                            }
                            catch (Exception ex)
                            {
                                // Ghi log / bỏ qua phần tử không hợp lệ
                                // Bạn có thể mở rộng để lưu danh sách lỗi
                                System.Diagnostics.Debug.WriteLine($"Error disallow join for {beam.Id}: {ex.Message}");
                            }

                            processed++;
                            // Cập nhật progress
                            int pct = (int)(processed * 100.0 / beams.Count);
                            pf.UpdateProgress(pct, processed, beams.Count);
                            Application.DoEvents(); // đơn giản để UI update
                        }

                        t.Commit();
                    }

                    pf.Close();
                }

                TaskDialog.Show("Done", "Đã đặt disallow join cho cả 2 đầu của tất cả dầm.\nSố dầm xử lý: " + beams.Count);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    // Simple WinForm for progress display
    public class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblInfo;
        private int totalCount;

        public ProgressForm(int total)
        {
            totalCount = total;
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Processing: Disallow Beam Joins";
            this.Width = 420;
            this.Height = 120;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;

            lblInfo = new Label()
            {
                Text = $"Processing 0 / {totalCount}",
                AutoSize = false,
                Width = 380,
                Height = 20,
                Top = 10,
                Left = 10
            };

            progressBar = new ProgressBar()
            {
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Width = 380,
                Height = 20,
                Top = 40,
                Left = 10
            };

            this.Controls.Add(lblInfo);
            this.Controls.Add(progressBar);
        }

        public void UpdateProgress(int percentage, int processed, int total)
        {
            if (percentage < 0) percentage = 0;
            if (percentage > 100) percentage = 100;
            progressBar.Value = percentage;
            lblInfo.Text = $"Processing {processed} / {total}  ({percentage} %)";
            // đảm bảo repaint nhanh
            progressBar.Refresh();
            lblInfo.Refresh();
        }
    }
}
