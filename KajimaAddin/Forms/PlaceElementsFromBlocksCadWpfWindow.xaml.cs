using System.Windows;
using Autodesk.Revit.UI;
using SKToolsAddins.Commands.PlaceElementsFromBlocksCad;
using SKToolsAddins.Utils;
using SKToolsAddins.ViewModel;
using RequestId = SKToolsAddins.Commands.PlaceElementsFromBlocksCad.RequestId;

namespace SKToolsAddins.Forms
{
    public partial class PlaceElementsFromBlocksCadWpfWindow : Window
    {
        private PlaceElementsFromBlocksCadRequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        public PlaceElementsFromBlocksCadWpfWindow(ExternalEvent exEvent,
            PlaceElementsFromBlocksCadRequestHandler handler,
            PlaceElementsFromBlocksCadViewModel viewModel)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;

            this.LoadViewFromUri("/KajimaRevitAddins;componenet/Forms/PlaceElementsFromBlocksCadWpfWindow.xaml");

            this.DataContext = viewModel;

            OkBtn.Click += OkBtn_Click;
            CancelBtn.Click += CancelBtn_Click;



        }
        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void MakeRequest(Commands.PlaceElementsFromBlocksCad.RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }
 
    }
}
