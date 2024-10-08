using System.Windows;
using Autodesk.Revit.UI;
using SKRevitAddins.Commands.PermissibleRangeFrame;
using SKRevitAddins.ViewModel;
using SKRevitAddins.Utils;
using RequestId = SKRevitAddins.Commands.PermissibleRangeFrame.RequestId;

namespace SKRevitAddins.Forms
{
    public partial class PermissibleRangeFrameWpfWindow : Window
    {
        private PermissibleRangeFrameRequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        public PermissibleRangeFrameWpfWindow(ExternalEvent exEvent,
            PermissibleRangeFrameRequestHandler handler,
            PermissibleRangeFrameViewModel viewModel)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;

            this.LoadViewFromUri("/KajimaRevitAddins;componenet/Forms/PermissibleRangeFrameWpfWindow.xaml");

            this.DataContext = viewModel;

            createBtn.Click += CreateBtn_Click;

            cancelBtn.Click += CancelBtn_Click;

            PermissibleRange.Click += PermissibleRange_Checked;

            AutoCreateSleeve.Click += AutoCreateSleeve_Checked;

            CreateReport.Click += CreateReport_Checked;

            SelectAllOption.Click += SelectAllOptionBtn_Click;

            DeSelectAllOption.Click += DeSelectAllOptionBtn_Click;
       
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void CreateBtn_Click(object sender, RoutedEventArgs e)
        {
            MakeRequest(RequestId.OK);
            this.Close();
        }
        private void PermissibleRange_Checked(object sender, RoutedEventArgs e)
        {
            AutoCreateSleeve.IsChecked = false;
            CreateReport.IsChecked = false;
        }
        private void AutoCreateSleeve_Checked(object sender, RoutedEventArgs e)
        {
            PermissibleRange.IsChecked = false;
            CreateReport.IsChecked = false;
        }
        private void CreateReport_Checked(object sender, RoutedEventArgs e)
        {
            PermissibleRange.IsChecked = false;
            AutoCreateSleeve.IsChecked = false;
        }
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
       
        private void MakeRequest(Commands.PermissibleRangeFrame.RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }
 
    }
}
