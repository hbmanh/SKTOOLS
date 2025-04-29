using System.Windows;
using Autodesk.Revit.UI;
using SKRevitAddins.Commands.CopySetOfFilterFromViewTemp;
using SKRevitAddins.ViewModel;
using SKRevitAddins.Utils;
using RequestId = SKRevitAddins.Commands.CopySetOfFilterFromViewTemp.RequestId;

namespace SKRevitAddins.Forms
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

            //this.LoadViewFromUri("/KajimaRevitAddins;componenet/Forms/CopySetOfFilterFromViewTempWpfWindow.xaml");

            this.DataContext = viewModel;

            createBtn.Click += CreateBtn_Click;

            cancelBtn.Click += CancelBtn_Click;

            AllCopyBOX.Click += AllCopy_Checked;

            PatternCopyBOX.Click += PatternCopy_Checked;

            CutSetCopyBOX.Click += CutSetCopy_Checked;

            AllSelFiltersBtn.Click += AllSelFiltersBtn_Click;

            NonSelFiltersBtn.Click += NonSelFiltersBtn_Click;

            AllSelViewTargetBtn.Click += AllSelViewTargetBtn_Click;

            NonSelViewTargetBtn.Click += NonSelViewTargetBtn_Click;
       
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
        private void AllCopy_Checked(object sender, RoutedEventArgs e)
        {
            PatternCopyBOX.IsChecked = false;
            CutSetCopyBOX.IsChecked = false;
        }
        private void PatternCopy_Checked(object sender, RoutedEventArgs e)
        {
            AllCopyBOX.IsChecked = false;
            CutSetCopyBOX.IsChecked = false;
        }
        private void CutSetCopy_Checked(object sender, RoutedEventArgs e)
        {
            AllCopyBOX.IsChecked = false;
            PatternCopyBOX.IsChecked = false;
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
        private void MakeRequest(Commands.CopySetOfFilterFromViewTemp.RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
        }
 
    }
}
