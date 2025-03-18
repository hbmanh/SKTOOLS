using System.Windows;
using SKRevitAddins.ViewModel;
using System.Data;
using System.Collections.Generic;

namespace SKRevitAddins.Forms
{
    public partial class PreviewWindow : Window
    {
        private ExportSchedulesToExcelViewModel _vm;
        private List<ExportSchedulesToExcelViewModel.PreviewTab> _originalTabs;

        public PreviewWindow(ExportSchedulesToExcelViewModel viewModel)
        {
            InitializeComponent();
            _vm = viewModel;
            DataContext = _vm;

            // Sao lưu dữ liệu ban đầu
            _originalTabs = new List<ExportSchedulesToExcelViewModel.PreviewTab>();
            foreach (var tab in _vm.ExcelPreviewTabs)
            {
                var copyTab = new ExportSchedulesToExcelViewModel.PreviewTab
                {
                    SheetName = tab.SheetName,
                    SheetData = tab.SheetData.Copy() // Copy DataTable
                };
                _originalTabs.Add(copyTab);
            }
        }

        private void ResetBtn_Click(object sender, RoutedEventArgs e)
        {
            _vm.ExcelPreviewTabs.Clear();
            foreach (var oldTab in _originalTabs)
            {
                var newTab = new ExportSchedulesToExcelViewModel.PreviewTab
                {
                    SheetName = oldTab.SheetName,
                    SheetData = oldTab.SheetData.Copy()
                };
                _vm.ExcelPreviewTabs.Add(newTab);
            }
        }

        private void UpdateBtn_Click(object sender, RoutedEventArgs e)
        {
            // Logic ghi ngược Revit (nếu muốn)
            MessageBox.Show("Data updated (demo).",
                "Update", MessageBoxButton.OK, MessageBoxImage.Information);

            this.Close();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    public static class DataTableExtensions
    {
        public static DataTable Copy(this DataTable dt)
        {
            // Copy DataTable
            return dt.Copy();
        }
    }
}
