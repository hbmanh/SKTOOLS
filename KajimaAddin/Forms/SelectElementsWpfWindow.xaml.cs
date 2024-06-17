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
using SKToolsAddins.Commands.SelectElements;
using SKToolsAddins.ViewModel;
using SelectionChangedEventArgs = System.Windows.Controls.SelectionChangedEventArgs;
using Window = System.Windows.Window;

namespace SKToolsAddins.Forms
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