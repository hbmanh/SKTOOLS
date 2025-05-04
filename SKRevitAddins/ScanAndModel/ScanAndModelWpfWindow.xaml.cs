using System.Windows;
using ScanAndModel.ViewModel;
using Autodesk.Revit.UI;
using SKRevitAddins.ScanAndModel;

namespace ScanAndModel.ScanAndModel
{
    public partial class ScanAndModelWpfWindow : Window
    {
        private ExternalEvent _exEvent;
        private ScanAndModelRequestHandler _handler;
        private ScanAndModelViewModel _vm;

        public ScanAndModelWpfWindow(
            ExternalEvent exEvent,
            ScanAndModelRequestHandler handler,
            ScanAndModelViewModel viewModel)
        {
            InitializeComponent();
            _exEvent = exEvent;
            _handler = handler;
            _vm = viewModel;
            DataContext = _vm;
        }

        private void AutoDetectBtn_Click(object sender, RoutedEventArgs e)
        {
            // Gửi request AutoDetectAndModel
            _handler.Request.Make(ScanAndModelRequestId.AutoDetectAndModel);
            _exEvent.Raise();
        }

        private void ZoomToPointBtn_Click(object sender, RoutedEventArgs e)
        {
            // Gửi request ZoomToPoint
            _handler.Request.Make(ScanAndModelRequestId.ZoomToPoint);
            _exEvent.Raise();
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
