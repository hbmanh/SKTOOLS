using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;

namespace SKRevitAddins.FindDWGNotUseAndDel
{
    public partial class FindDWGNotUseAndDelWpfWindow : Window
    {
        private ExternalEvent _exEvent;
        private FindDWGNotUseAndDelRequestHandler _handler;
        private FindDWGNotUseAndDelViewModel _vm;

        public FindDWGNotUseAndDelWpfWindow(
            ExternalEvent exEvent,
            FindDWGNotUseAndDelRequestHandler handler,
            FindDWGNotUseAndDelViewModel viewModel)
        {
            InitializeComponent();

            _exEvent = exEvent;
            _handler = handler;
            _vm = viewModel;
            this.DataContext = _vm;
        }

        // Lấy các dòng được chọn, lưu vào _vm.SelectedDWGs
        private void viewSetDg_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Cast SelectedItems => DwgItem
            var selectedDwgItems = viewSetDg.SelectedItems
                .OfType<FindDWGNotUseAndDelViewModel.DwgItem>()
                .ToList();

            _vm.SelectedDWGs = selectedDwgItems;
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string keyword = SearchTextBox.Text?.Trim();
            _vm.FilterDWGs(keyword);
        }

        private void DeleteBtn_Click(object sender, RoutedEventArgs e)
        {
            // All or Selected
            _vm.IsDeleteAll = AllRadioBtn.IsChecked == true;

            _handler.Request.Make(RequestId.Delete);
            _exEvent.Raise();

            // (Tuỳ ý) Đóng cửa sổ sau khi xoá
            // this.Close();
        }

        private void OpenViewBtn_Click(object sender, RoutedEventArgs e)
        {
            _handler.Request.Make(RequestId.OpenView);
            _exEvent.Raise();
        }

        private void ExportTableBtn_Click(object sender, RoutedEventArgs e)
        {
            _handler.Request.Make(RequestId.Export);
            _exEvent.Raise();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
