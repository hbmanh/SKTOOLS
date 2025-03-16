using System.Collections.Generic;
using System.Windows;
using SKRevitAddins.ViewModel;

namespace SKRevitAddins.Forms
{
    public partial class PreviewWindow : Window
    {
        private List<List<string>> _originalExcelData;
        private ExportSchedulesToExcelViewModel _viewModel;

        public PreviewWindow(ExportSchedulesToExcelViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;

            // Lưu bản sao dữ liệu ban đầu để có thể Reset
            _originalExcelData = new List<List<string>>(_viewModel.ExcelPreviewData);

            // Hiển thị dữ liệu trong DataGrid
            ExcelPreviewDataGrid.ItemsSource = _viewModel.ExcelPreviewData;
        }

        // Sự kiện Reset: Quay lại dữ liệu gốc
        private void ResetBtn_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ExcelPreviewData.Clear();
            foreach (var row in _originalExcelData)
            {
                // copy từng row
                _viewModel.ExcelPreviewData.Add(new List<string>(row));
            }
        }

        // Sự kiện Update: Cập nhật dữ liệu đã chỉnh sửa vào Revit model (nếu cần)
        private void UpdateBtn_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Data updated successfully (logic not shown).",
                "Update", MessageBoxButton.OK, MessageBoxImage.Information);

            // Đóng cửa sổ Preview sau khi cập nhật
            this.Close();
        }

        // Sự kiện Cancel: Đóng cửa sổ Preview mà không lưu thay đổi
        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
