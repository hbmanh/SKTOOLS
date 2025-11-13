using System.Windows;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;

namespace SKRevitAddins.CopySetOfFilterFromViewTemp
{
    public partial class CopySetOfFilterFromViewTempWpfWindow : Window
    {
        private readonly CopySetOfFilterFromViewTempRequestHandler m_Handler;
        private readonly ExternalEvent m_ExEvent;

        public CopySetOfFilterFromViewTempWpfWindow(ExternalEvent exEvent,
            CopySetOfFilterFromViewTempRequestHandler handler,
            CopySetOfFilterFromViewTempViewModel viewModel)
        {
            InitializeComponent();
            m_Handler = handler;
            m_ExEvent = exEvent;
            DataContext = viewModel;

            // Try to load the logo dynamically
            LogoHelper.TryLoadLogo(LogoImage);
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e) => Close();

        private void CreateBtn_Click(object sender, RoutedEventArgs e)
        {
            MakeRequest(RequestId.OK);
            Close();
        }

        private void AllSelFiltersBtn_Click(object sender, RoutedEventArgs e) => Filterlb.SelectAll();

        private void NonSelFiltersBtn_Click(object sender, RoutedEventArgs e) => Filterlb.UnselectAll();

        private void AllSelViewTargetBtn_Click(object sender, RoutedEventArgs e) => ViewTargetlb.SelectAll();

        private void NonSelViewTargetBtn_Click(object sender, RoutedEventArgs e) => ViewTargetlb.UnselectAll();

        private void MakeRequest(RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }
    }
}
