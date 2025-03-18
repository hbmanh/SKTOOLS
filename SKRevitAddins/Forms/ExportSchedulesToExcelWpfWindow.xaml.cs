using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;
using SKRevitAddins.Commands.ExportSchedulesToExcel;
using SKRevitAddins.ViewModel;

namespace SKRevitAddins.Forms
{
    public partial class ExportSchedulesToExcelWpfWindow : Window
    {
        private ExternalEvent _exEvent;
        private ExportSchedulesToExcelRequestHandler _handler;
        private ExportSchedulesToExcelViewModel _vm;

        public ExportSchedulesToExcelWpfWindow(
            ExternalEvent exEvent,
            ExportSchedulesToExcelRequestHandler handler,
            ExportSchedulesToExcelViewModel viewModel)
        {
            InitializeComponent();
            _exEvent = exEvent;
            _handler = handler;
            _vm = viewModel;
            DataContext = _vm;
        }

        // Search Document
        private void DocSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string keyword = DocSearchTextBox.Text?.Trim();
            _vm.FilterDocumentByKeyword(keyword);
        }

        // Search Schedule
        private void SchedSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string keyword = SchedSearchTextBox.Text?.Trim();
            _vm.FilterScheduleByKeyword(keyword);
        }

        // CheckBox => cập nhật SelectedSchedules
        private void ScheduleCheckbox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb && cb.DataContext is ExportSchedulesToExcelViewModel.ScheduleItem item)
            {
                if (item.IsSelected)
                {
                    if (!_vm.SelectedSchedules.Contains(item))
                        _vm.SelectedSchedules.Add(item);
                }
                else
                {
                    if (_vm.SelectedSchedules.Contains(item))
                        _vm.SelectedSchedules.Remove(item);
                }
            }
        }

        // Select All
        private void SelectAllBtn_Click(object sender, RoutedEventArgs e)
        {
            foreach (var s in _vm.FilteredSchedules)
            {
                if (!s.IsSelected)
                {
                    s.IsSelected = true;
                    if (!_vm.SelectedSchedules.Contains(s))
                        _vm.SelectedSchedules.Add(s);
                }
            }
        }

        // Deselect All
        private void DeselectAllBtn_Click(object sender, RoutedEventArgs e)
        {
            foreach (var s in _vm.FilteredSchedules)
            {
                s.IsSelected = false;
            }
            _vm.SelectedSchedules.Clear();
        }

        // Preview
        private void PreviewBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!_vm.SelectedSchedules.Any())
            {
                MessageBox.Show("No schedules selected to preview.",
                                "Preview", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Tạo tab
            _vm.LoadPreviewTabsFromSelectedSchedules();

            // Mở cửa sổ Preview
            PreviewWindow pwin = new PreviewWindow(_vm);
            pwin.ShowDialog();
        }

        // Export
        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!_vm.SelectedSchedules.Any())
            {
                MessageBox.Show("No schedules selected to export.",
                                "Export", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Gọi ExternalEvent => Export
            _handler.Request.Make(RequestId.Export);
            _exEvent.Raise();
        }
    }
}
