﻿using System.Linq;
using System.Windows;
using System.Windows.Controls;
using SKRevitAddins.ViewModel;
using SKRevitAddins.Commands.ExportSchedulesToExcel;
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

        private void ScheduleCheckbox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb &&
                cb.DataContext is ExportSchedulesToExcelViewModel.ScheduleItem item)
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

        private void DeselectAllBtn_Click(object sender, RoutedEventArgs e)
        {
            foreach (var s in _vm.FilteredSchedules)
            {
                s.IsSelected = false;
            }
            _vm.SelectedSchedules.Clear();
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!_vm.SelectedSchedules.Any())
            {
                _vm.ExportStatusMessage = "No schedules selected to export.";
                return;
            }

            // Gọi ExternalEvent => Export
            _handler.Request.Make(RequestId.Export);
            _exEvent.Raise();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            _vm.IsCancelled = true;
            this.Close();
        }
    }
}
