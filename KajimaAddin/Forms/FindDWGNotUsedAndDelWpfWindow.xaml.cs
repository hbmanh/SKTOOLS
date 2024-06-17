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
using SKToolsAddins.Commands.FindDWGNotUseAndDel;
using SKToolsAddins.ViewModel;
using RequestId = SKToolsAddins.Commands.FindDWGNotUseAndDel.RequestId;
using SelectionChangedEventArgs = System.Windows.Controls.SelectionChangedEventArgs;
using Window = System.Windows.Window;

namespace SKToolsAddins.Forms
{
    public partial class FindDWGNotUseAndDelWpfWindow : Window
    {
    private FindDWGNotUsedAndDelRequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        private FindDWGNotUsedAndDelViewModel viewModel;
        public FindDWGNotUseAndDelWpfWindow(ExternalEvent exEvent,
            FindDWGNotUsedAndDelRequestHandler handler,
            FindDWGNotUsedAndDelViewModel viewModel)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;

            this.LoadViewFromUri("/KajimaRevitAddins;componenet/Forms/FindDWGNotUseAndDelWpfWindow.xaml");

            this.DataContext = viewModel;
            this.viewModel = viewModel;


            AllBtn.Click += AllBtn_Click;

            NoneBtn.Click += NoneBtn_Click;

            DelBtn.Click += DelBtn_Click;

            ExportBtn.Click += ExportBtn_Click;

            CancelBtn.Click += CancelBtn_Click;



        }

        private void AllBtn_Click(object sender, RoutedEventArgs e)
        {
        }

        private void NoneBtn_Click(object sender, RoutedEventArgs e)
        {
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void DelBtn_Click(object sender, RoutedEventArgs e)
        {
            MakeRequest(RequestId.OK);
            this.Close();
        }
        private void MakeRequest(Commands.FindDWGNotUseAndDel.RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }
        
    }
}
