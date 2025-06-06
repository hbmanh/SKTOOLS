﻿using System.Windows;
using Autodesk.Revit.UI;

namespace SKRevitAddins.PermissibleRangeFrame
{
    public partial class PermissibleRangeFrameWpfWindow : Window
    {
        private PermissibleRangeFrameRequestHandler _handler;
        private ExternalEvent _exEvent;

        public PermissibleRangeFrameWpfWindow(
            ExternalEvent exEvent,
            PermissibleRangeFrameRequestHandler handler,
            PermissibleRangeFrameViewModel viewModel)
        {
            InitializeComponent();

            _handler = handler;
            _exEvent = exEvent;
            this.DataContext = viewModel;

            createBtn.Click += CreateBtn_Click;
            cancelBtn.Click += CancelBtn_Click;
            PermissibleRange.Click += PermissibleRange_Checked;
            AutoCreateSleeve.Click += AutoCreateSleeve_Checked;
            CreateReport.Click += CreateReport_Checked;
            SelectAllOption.Click += SelectAllOptionBtn_Click;
            DeSelectAllOption.Click += DeSelectAllOptionBtn_Click;
            previewBtn.Click += PreviewBtn_Click;
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e) => this.Close();

        private void CreateBtn_Click(object sender, RoutedEventArgs e)
        {
            MakeRequest(RequestId.OK);
            this.Close();
        }

        private void PermissibleRange_Checked(object sender, RoutedEventArgs e) { }
        private void AutoCreateSleeve_Checked(object sender, RoutedEventArgs e) { }
        private void CreateReport_Checked(object sender, RoutedEventArgs e) { }

        private void SelectAllOptionBtn_Click(object sender, RoutedEventArgs e)
        {
            PermissibleRange.IsChecked = true;
            AutoCreateSleeve.IsChecked = true;
            CreateReport.IsChecked = true;
        }

        private void DeSelectAllOptionBtn_Click(object sender, RoutedEventArgs e)
        {
            PermissibleRange.IsChecked = false;
            AutoCreateSleeve.IsChecked = false;
            CreateReport.IsChecked = false;
        }

        private void MakeRequest(RequestId request)
        {
            _handler.Request.Make(request);
            _exEvent.Raise();
        }

        // Nút Preview -> mở cửa sổ PreviewReportWindow
        private void PreviewBtn_Click(object sender, RoutedEventArgs e)
        {
            if (this.DataContext is PermissibleRangeFrameViewModel vm)
            {
                var previewWindow = new PreviewReportWindow(vm.ErrorMessages, vm.UiApp);
                previewWindow.Owner = this;
                previewWindow.ShowDialog();
            }
        }
    }
}
