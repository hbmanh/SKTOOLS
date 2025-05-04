using System.Windows;
using Autodesk.Revit.UI;
using SKRevitAddins.CreateSpace;
using SKRevitAddins.ViewModel;
using SKRevitAddins.Utils;

namespace SKRevitAddins.Forms
{
    public partial class CreateSpaceWpfWindow : Window
    {
        private CreateSpaceRequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        public CreateSpaceWpfWindow(ExternalEvent exEvent,
            CreateSpaceRequestHandler handler,
            CreateSpaceViewModel createSpaceViewModel)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;

            this.LoadViewFromUri("/KajimaRevitAddins;componenet/Forms/CreateSpaceWpfWindow.xaml");

            this.DataContext = createSpaceViewModel;

            createBtn.Click += CreateBtn_Click;

            cancelBtn.Click += CancelBtn_Click;

            selectAllBtn.Click += SelectAllBtn_Click;

            nonBtn.Click += NoneBtn_Click;

            deleteSpaceBtn.Click += DeleteSpaceBtn_Click;

            TagPlacementBOX.Click += TagPlacementBOX_Checked;

            SpaceOffsetBOX.Click += SpaceOffsetBOX_Checked;

            NameNumberBOX.Click += NameNumberBOX_Checked;

        }

        private void DeleteSpaceBtn_Click(object sender, RoutedEventArgs e)
        {
            MakeRequest(RequestId.DeleteSpace);
            this.Close();
        }
        private void SelectAllBtn_Click(object sender, RoutedEventArgs e)
        {
            lbViewName.SelectAll();
        }
        private void NoneBtn_Click(object sender, RoutedEventArgs e)
        {
            lbViewName.UnselectAll();
        }
        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void CreateBtn_Click(object sender, RoutedEventArgs e)
        {
            MakeRequest(RequestId.CreateSpace);
            this.Close();
        }

        private void MakeRequest(RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }

        private void TagPlacementBOX_Checked(object sender, RoutedEventArgs e)
        {
        }
        private void SpaceOffsetBOX_Checked(object sender, RoutedEventArgs e)
        {
        }

        private void NameNumberBOX_Checked(object sender, RoutedEventArgs e)
        {
        }
    }
}
