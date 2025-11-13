using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;

namespace SKRevitAddins.ExportSchedulesToExcel
{
    public partial class ExportSchedulesToExcelWpfWindow : Window
    {
        private readonly ExternalEvent _exEvent;
        private readonly ExportSchedulesToExcelRequestHandler _handler;
        private readonly ExportSchedulesToExcelViewModel _vm;

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

            LogoHelper.TryLoadLogo(LogoImage);
        }

        private void DocSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string keyword = DocSearchTextBox.Text?.Trim();
            _vm.FilterDocumentByKeyword(keyword);
        }

        private void SchedSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string keyword = SchedSearchTextBox.Text?.Trim();
            _vm.FilterScheduleByKeyword(keyword);
        }

        private void SchedulesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ListBox lb)
            {
                _vm.SelectedSchedules.Clear();
                foreach (var item in lb.SelectedItems)
                {
                    if (item is ExportSchedulesToExcelViewModel.ScheduleItem schedule)
                        _vm.SelectedSchedules.Add(schedule);
                }
            }
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!_vm.SelectedSchedules.Any())
            {
                _vm.ExportStatusMessage = "No schedules selected to export.";
                return;
            }

            _handler.Request.Make(RequestId.Export);
            _exEvent.Raise();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            _vm.IsCancelled = true;
            Close();
        }
    }
}
