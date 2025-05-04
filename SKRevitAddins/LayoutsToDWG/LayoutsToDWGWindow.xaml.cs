using Autodesk.Revit.UI;
using System.Windows;

namespace SKRevitAddins.LayoutsToDWG
{
    public partial class LayoutsToDWGWindow : Window
    {
        private readonly ExternalEvent _externalEvent;
        private readonly LayoutsToDWGRequestHandler _handler;
        private readonly LayoutsToDWGViewModel _viewModel;

        public LayoutsToDWGWindow(ExternalEvent externalEvent,
            LayoutsToDWGRequestHandler handler,
            LayoutsToDWGViewModel viewModel)
        {
            InitializeComponent();

            _externalEvent = externalEvent;
            _handler = handler;
            _viewModel = viewModel;

            // Cho handler nhìn thấy ViewModel
            _handler.ViewModel = _viewModel;

            DataContext = _viewModel;
        }

        private void OpenSheetSelectionDialog_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SheetSelectionWindow(_viewModel);
            dlg.ShowDialog();
        }

        private void Export_Click(object sender, RoutedEventArgs e) => _externalEvent.Raise();
        private void Cancel_Click(object sender, RoutedEventArgs e) => Close();
    }
}