using System.Windows;
using Autodesk.Revit.UI;
using SKToolsAddins.Commands.PlaceElementsFromBlocksCad;
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

            DataContext = viewModel;

            OkBtn.Click += OkBtn_Click;
            CancelBtn.Click += CancelBtn_Click;
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            MakeRequest(RequestId.OK);
            this.Close();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void MakeRequest(RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }
    }
}