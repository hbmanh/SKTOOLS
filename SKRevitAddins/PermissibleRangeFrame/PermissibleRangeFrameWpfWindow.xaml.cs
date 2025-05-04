using System.Windows;
using Autodesk.Revit.UI;
using SKRevitAddins.PermissibleRangeFrame;
using SKRevitAddins.ViewModel;

namespace SKRevitAddins.PermissibleRangeFrame
{
    public partial class PermissibleRangeFrameWpfWindow : Window
    {
        private PermissibleRangeFrameRequestHandler m_Handler;
        private ExternalEvent m_ExEvent;

        public PermissibleRangeFrameWpfWindow(
            ExternalEvent exEvent,
            PermissibleRangeFrameRequestHandler handler,
            PermissibleRangeFrameViewModel viewModel)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;
            this.DataContext = viewModel;

            createBtn.Click += CreateBtn_Click;
            cancelBtn.Click += CancelBtn_Click;
            PermissibleRange.Click += PermissibleRange_Checked;
            AutoCreateSleeve.Click += AutoCreateSleeve_Checked;
            CreateReport.Click += CreateReport_Checked;
            SelectAllOption.Click += SelectAllOptionBtn_Click;
            DeSelectAllOption.Click += DeSelectAllOptionBtn_Click;
            previewBtn.Click += PreviewBtn_Click; // Thêm sự kiện cho nút Preview
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
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }

        // Nút Preview -> mở cửa sổ PreviewReportWindow
        private void PreviewBtn_Click(object sender, RoutedEventArgs e)
        {
            // Lấy ViewModel để truyền ErrorMessages và UiApp
            if (this.DataContext is PermissibleRangeFrameViewModel vm)
            {
                // Mở cửa sổ Preview
                var previewWindow = new PreviewReportWindow(vm.ErrorMessages, vm.UiApp);
                previewWindow.Owner = this; // Đặt owner cho window
                previewWindow.ShowDialog();
            }
        }
    }
}
