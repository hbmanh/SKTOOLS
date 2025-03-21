using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.Forms
{
    public partial class PreviewReportWindow : Window
    {
        private Dictionary<ElementId, HashSet<string>> _errorMessages;
        private UIApplication _uiApp;

        public PreviewReportWindow(Dictionary<ElementId, HashSet<string>> errorMessages, UIApplication uiApp)
        {
            InitializeComponent();
            _errorMessages = errorMessages;
            _uiApp = uiApp;

            // Chuyển dictionary sang list hiển thị trong DataGrid
            var dataList = new List<ErrorItem>();
            foreach (var kvp in errorMessages)
            {
                ElementId eid = kvp.Key;
                var doc = _uiApp.ActiveUIDocument.Document;
                var element = doc.GetElement(eid);
                string elementName = element?.Name ?? "Unknown Element";
                string msg = string.Join("; ", kvp.Value);

                dataList.Add(new ErrorItem
                {
                    ElementId = eid.IntegerValue,
                    ElementName = elementName,
                    ErrorDescription = msg
                });
            }

            ErrorsDataGrid.ItemsSource = dataList;
        }

        // Khi nhấn nút Show -> zoom tới element
        private void ShowBtn_Click(object sender, RoutedEventArgs e)
        {
            if (ErrorsDataGrid.SelectedItem is ErrorItem selected)
            {
                var doc = _uiApp.ActiveUIDocument.Document;
                var eid = new ElementId(selected.ElementId);
                if (eid != ElementId.InvalidElementId)
                {
                    _uiApp.ActiveUIDocument.ShowElements(eid);
                }
            }
        }

        // Khi nhấn nút Export -> Xuất file CSV (minh hoạ)
        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            var data = ErrorsDataGrid.ItemsSource as List<ErrorItem>;
            if (data == null || !data.Any()) return;

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV file (*.csv)|*.csv",
                FileName = "ErrorReport.csv"
            };
            if (dlg.ShowDialog() == true)
            {
                var lines = new List<string>();
                lines.Add("ElementId,ElementName,ErrorDescription");
                foreach (var item in data)
                {
                    lines.Add($"{item.ElementId},{item.ElementName},\"{item.ErrorDescription}\"");
                }
                System.IO.File.WriteAllLines(dlg.FileName, lines);
                MessageBox.Show("Exported to " + dlg.FileName);
            }
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    // Lớp chứa thông tin hiển thị trên DataGrid
    public class ErrorItem
    {
        public int ElementId { get; set; }
        public string ElementName { get; set; }
        public string ErrorDescription { get; set; }
    }
}
