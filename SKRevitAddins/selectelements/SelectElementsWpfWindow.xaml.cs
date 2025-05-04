using System.Windows;
using Autodesk.Revit.UI;
using SKRevitAddins.SelectElements;
using SKRevitAddins.ViewModel;
using SKRevitAddins.Utils;
using Window = System.Windows.Window;

namespace SKRevitAddins.SelectElements
{
    public partial class SelectElementsWpfWindow : Window
    {
        private SelectElementsRequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        private SelectElementsViewModel viewModel;
        public SelectElementsWpfWindow(ExternalEvent exEvent,
            SelectElementsRequestHandler handler,
            SelectElementsViewModel viewModel)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;

            this.LoadViewFromUri("/KajimaRevitAddins;componenet/Forms/SelectElementsWpfWindow.xaml");

            this.DataContext = viewModel;
            this.viewModel = viewModel;




        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void ReviewBtn_Click(object sender, RoutedEventArgs e)
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