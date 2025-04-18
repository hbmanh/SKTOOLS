using System.Windows;
using Autodesk.Revit.UI;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    public partial class LayoutsToDWGWindow : Window
    {
        readonly ExternalEvent _externalEvent;
        readonly LayoutsToDWGRequestHandler _handler;
        readonly LayoutsToDWGViewModel _viewModel;

        public LayoutsToDWGWindow(ExternalEvent ev,
                                  LayoutsToDWGRequestHandler handler,
                                  LayoutsToDWGViewModel vm)
        {
            InitializeComponent();
            _externalEvent = ev;
            _handler = handler;
            _viewModel = vm;
            DataContext = vm;
        }

        private void OpenSheetSelectionDialog_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SheetSelectionWindow(_viewModel);
            dlg.ShowDialog();
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            _externalEvent.Raise();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
