using System.Windows;
using Autodesk.Revit.UI;
using SKRevitAddins.PlaceElementsFromBlocksCad;
using SKRevitAddins.ViewModel;
using Commands_PlaceElementsFromBlocksCad_RequestId = SKRevitAddins.PlaceElementsFromBlocksCad.RequestId;
using PlaceElementsFromBlocksCad_RequestId = SKRevitAddins.PlaceElementsFromBlocksCad.RequestId;
using RequestId = SKRevitAddins.PlaceElementsFromBlocksCad.RequestId;

namespace SKRevitAddins.PlaceElementsFromBlocksCad
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

        private void MakeRequest(Commands_PlaceElementsFromBlocksCad_RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }
    }
}