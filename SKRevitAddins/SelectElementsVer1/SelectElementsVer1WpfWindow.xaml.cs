using System.Windows;
using Autodesk.Revit.UI;
using SKRevitAddins.SelectElementsVer1;
using SKRevitAddins.ViewModel;
using SKRevitAddins.Utils;
using Window = System.Windows.Window;

namespace SKRevitAddins.SelectElementsVer1
{
    public partial class SelectElementsVer1WpfWindow : Window
    {
    private SelectElementsVer1RequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        private SelectElementsVer1ViewModel viewModel;
        public SelectElementsVer1WpfWindow(ExternalEvent exEvent,
            SelectElementsVer1RequestHandler handler,
            SelectElementsVer1ViewModel viewModel)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;

            this.LoadViewFromUri("/KajimaRevitAddins;componenet/Forms/SelectElementsWpfWindow.xaml");

            this.DataContext = viewModel;
            this.viewModel = viewModel;

            valSetBtn.Click += ValSetBtn_Click;

            NumberingBtn.Click += Numbering_Click;

            CancelBtn.Click += CancelBtn_Click;

        }

        private void ValSetBtn_Click(object sender, RoutedEventArgs e)
        {
        }
        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void Numbering_Click(object sender, RoutedEventArgs e)
        {
            MakeRequest(RequestId.OK);
            this.Close();
        }
        private void MakeRequest(RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }
        
    }
}
