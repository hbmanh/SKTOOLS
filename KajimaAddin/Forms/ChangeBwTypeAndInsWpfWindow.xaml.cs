using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKToolsAddins.Utils;
using SKToolsAddins.Commands.ChangeBwTypeAndIns;
using SKToolsAddins.ViewModel;
using RequestId = SKToolsAddins.Commands.ChangeBwTypeAndIns.RequestId;
using SelectionChangedEventArgs = System.Windows.Controls.SelectionChangedEventArgs;
using Window = System.Windows.Window;

namespace SKToolsAddins.Forms
{
    public partial class ChangeBwTypeAndInsWpfWindow : Window
    {
    private ChangeBwTypeAndInsRequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        private ChangeBwTypeAndInsViewModel viewModel;
        public ChangeBwTypeAndInsWpfWindow(ExternalEvent exEvent,
            ChangeBwTypeAndInsRequestHandler handler,
            ChangeBwTypeAndInsViewModel viewModel)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;

            this.LoadViewFromUri("/KajimaRevitAddins;componenet/Forms/ChangeBwTypeAndInsWpfWindow.xaml");

            this.DataContext = viewModel;
            this.viewModel = viewModel;

            CreateBtn.Click += CreateBtn_Click;

            CancelBtn.Click += CancelBtn_Click;

            InsToTypFamBtn.Click += InsToTypFamBtn_Click;

            TypeToInsFamBtn.Click += TypeToInsFamBtn_Click;

            InsToTypProBtn.Click += InsToTypProBtn_Click;

            TypeToInsProBtn .Click += TypeToInsProBtn_Click;

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
        private void InsToTypFamBtn_Click(object sender, RoutedEventArgs e)
        {
           
        }
        private void TypeToInsFamBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void InsToTypProBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void TypeToInsProBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        private void MakeRequest(Commands.ChangeBwTypeAndIns.RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }

        }
}
