using System.Windows;
using Autodesk.Revit.UI;
using SKRevitAddins.Commands.DeleteTypeOfTextNotesDontUse;
using SKRevitAddins.ViewModel;
using SKRevitAddins.Utils;
using Commands_DeleteTypeOfTextNotesDontUse_RequestId = SKRevitAddins.Commands.DeleteTypeOfTextNotesDontUse.RequestId;
using DeleteTypeOfTextNotesDontUse_RequestId = SKRevitAddins.Commands.DeleteTypeOfTextNotesDontUse.RequestId;
using RequestId = SKRevitAddins.Commands.DeleteTypeOfTextNotesDontUse.RequestId;

namespace SKRevitAddins.Forms
{
    public partial class DeleteTypeOfTextNotesDontUseWpfWindow : Window
    {
        private DeleteTypeOfTextNotesDontUseRequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        public DeleteTypeOfTextNotesDontUseWpfWindow(ExternalEvent exEvent,
            DeleteTypeOfTextNotesDontUseRequestHandler handler,
            DeleteTypeOfTextNotesDontUseViewModel viewModel)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;

            this.LoadViewFromUri("/KajimaRevitAddins;componenet/Forms/CopySetOfFilterFromViewTempWpfWindow.xaml");

            this.DataContext = viewModel;

            ReviewBtn.Click += ReviewBtn_Click;

            CancelBtn.Click += CancelBtn_Click;

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
        private void MakeRequest(Commands_DeleteTypeOfTextNotesDontUse_RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }
 
    }
}
