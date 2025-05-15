using System.Windows;
using Autodesk.Revit.UI;

namespace SKRevitAddins.CopySetOfFilterFromViewTemp
{
    public partial class CopySetOfFilterFromViewTempWpfWindow : Window
    {
        private CopySetOfFilterFromViewTempRequestHandler m_Handler;
        private ExternalEvent m_ExEvent;

        public CopySetOfFilterFromViewTempWpfWindow(ExternalEvent exEvent,
            CopySetOfFilterFromViewTempRequestHandler handler,
            CopySetOfFilterFromViewTempViewModel viewModel)
        {
            InitializeComponent();
            m_Handler = handler;
            m_ExEvent = exEvent;
            this.DataContext = viewModel;
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

        private void AllSelFiltersBtn_Click(object sender, RoutedEventArgs e)
        {
            Filterlb.SelectAll();
        }

        private void NonSelFiltersBtn_Click(object sender, RoutedEventArgs e)
        {
            Filterlb.UnselectAll();
        }

        private void AllSelViewTargetBtn_Click(object sender, RoutedEventArgs e)
        {
            ViewTargetlb.SelectAll();
        }

        private void NonSelViewTargetBtn_Click(object sender, RoutedEventArgs e)
        {
            ViewTargetlb.UnselectAll();
        }

        private void MakeRequest(RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }
    }
}