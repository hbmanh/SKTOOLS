using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Autodesk.Revit.UI;

namespace SKRevitAddins.Commands.DWGExport
{
    public partial class DWGExportWpfWindow : Window
    {
        private readonly ExternalEvent _ev;
        private readonly DWGExportRequestHandler _handler;
        private readonly DWGExportViewModel _vm;

        private readonly DispatcherTimer _debounce =
            new DispatcherTimer { Interval = System.TimeSpan.FromMilliseconds(300) };

        public DWGExportWpfWindow(ExternalEvent ev,
                                  DWGExportRequestHandler handler,
                                  DWGExportViewModel vm)
        {
            InitializeComponent();
            _ev = ev;
            _handler = handler;
            _vm = vm;
            DataContext = vm;

            _debounce.Tick += (_, __) =>
            {
                _debounce.Stop();
                _vm.RefreshCategoriesFromSelectedSheets();
            };
        }

        private void SheetsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _vm.SelectedSheets.Clear();
            foreach (var it in ((ListBox)sender).SelectedItems)
                _vm.SelectedSheets.Add((SheetItem)it);

            _debounce.Stop();
            _debounce.Start();
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            _handler.Request.Make(LayerExportRequestId.Export);
            _ev.Raise();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
