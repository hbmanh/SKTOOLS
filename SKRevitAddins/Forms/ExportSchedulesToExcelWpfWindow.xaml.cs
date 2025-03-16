using System.Linq;
using System.Windows;
using System.Windows.Controls;
using SKRevitAddins.Commands.ExportSchedulesToExcel;
using SKRevitAddins.ViewModel;
using System.Collections.Generic;
using Autodesk.Revit.UI;

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

        // 1. Sự kiện click CheckBox => Cập nhật SelectedSchedules
        private void ScheduleCheckbox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb && cb.DataContext is ExportSchedulesToExcelViewModel.ScheduleItem item)
            {
                // Nếu checkbox được check => thêm vào SelectedSchedules
                if (item.IsSelected)
                {
                    if (!_vm.SelectedSchedules.Contains(item))
                    {
                        _vm.SelectedSchedules.Add(item);
                    }
                }
                else
                {
                    // Nếu uncheck => gỡ khỏi SelectedSchedules
                    if (_vm.SelectedSchedules.Contains(item))
                    {
                        _vm.SelectedSchedules.Remove(item);
                    }
                }
            }
        }

        // 2. Select All => đánh dấu IsSelected = true cho tất cả
        private void SelectAllBtn_Click(object sender, RoutedEventArgs e)
        {
            foreach (var sched in _vm.AvailableSchedules)
            {
                if (!sched.IsSelected)
                {
                    sched.IsSelected = true;
                    if (!_vm.SelectedSchedules.Contains(sched))
                    {
                        _vm.SelectedSchedules.Add(sched);
                    }
                }
            }
            // Làm mới UI (nếu cần)
            SchedulesListBox.Items.Refresh();
        }

        // 3. Deselect All => đánh dấu IsSelected = false cho tất cả
        private void DeselectAllBtn_Click(object sender, RoutedEventArgs e)
        {
            foreach (var sched in _vm.AvailableSchedules)
            {
                sched.IsSelected = false;
            }
            _vm.SelectedSchedules.Clear();
            SchedulesListBox.Items.Refresh();
        }

        // 4. Preview => Mở cửa sổ preview
        private void PreviewBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!_vm.SelectedSchedules.Any())
            {
                MessageBox.Show("Please select at least one schedule to preview.",
                                "Preview", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            PreviewWindow previewWindow = new PreviewWindow(_vm);
            previewWindow.ShowDialog();
        }

        // 5. Export Without Preview
        private void ExportWithoutPreviewBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_vm.SelectedSchedules.Any())
            {
                _handler.Request.Make(RequestId.Export);
                _exEvent.Raise();
            }
            else
            {
                MessageBox.Show("Please select at least one schedule.", "Export", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
