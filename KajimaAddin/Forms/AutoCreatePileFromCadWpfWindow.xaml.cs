using System.Windows;
using Autodesk.Revit.UI;
using SKToolsAddins.Utils;
using SKToolsAddins.ViewModel;
using RequestId = SKToolsAddins.Commands.AutoCreatePileFromCad.RequestId;
using Window = System.Windows.Window;
using SKToolsAddins.Commands.AutoCreatePileFromCad;

namespace SKToolsAddins.Forms
{
    public partial class AutoCreatePileFromCadWpfWindow : Window
    {
        private AutoCreatePileFromCadRequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        private AutoCreatePileFromCadViewModel ViewModel;
        public AutoCreatePileFromCadWpfWindow(ExternalEvent exEvent,
            AutoCreatePileFromCadRequestHandler handler,
            AutoCreatePileFromCadViewModel viewModel)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;

            this.LoadViewFromUri("/KajimaRevitAddins;componenet/Forms/ChangeBwTypeAndInsWpfWindow.xaml");

            this.DataContext = viewModel;
            this.ViewModel = viewModel;

            OkBtn.Click += OkBtn_Click;

            CancelBtn.Click += CancelBtn_Click;

        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void OkBtn_Click(object sender, RoutedEventArgs e)
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
