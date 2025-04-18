using System.Linq;
using System.Windows;
using Autodesk.Revit.UI;

namespace LayoutsToDWG.Commands
{
    public partial class LayoutsToDWGWindow : Window
    {
        readonly ExternalEvent _ev;
        readonly LayoutsToDWGRequestHandler _handler;
        readonly LayoutsToDWGViewModel _vm;

        public LayoutsToDWGWindow(ExternalEvent ev,
                                  LayoutsToDWGRequestHandler handler,
                                  LayoutsToDWGViewModel vm)
        {
            InitializeComponent();
            _ev = ev; _handler = handler; _vm = vm;
            DataContext = vm;
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            if (!_vm.IsExportFolderValid)
            {
                MessageBox.Show("Chọn thư mục hợp lệ trước khi xuất.",
                                "Lỗi đường dẫn", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (_vm.SelectedSheets.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất một sheet.",
                                "Export DWG", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            _vm.Request.Make(LayoutExportRequestId.Export);
            _ev.Raise();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e) =>
            Close();

        private void OpenSheetSelectionDialog_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SheetSelectionWindow(_vm);
            if (dialog.ShowDialog() == true)
            {
                _vm.SelectedSheets.Clear();
                foreach (var s in dialog.SelectedSheets)
                    _vm.SelectedSheets.Add(s);
            }
        }
    }

    public partial class SheetSelectionWindow : Window
    {
        readonly LayoutsToDWGViewModel _vm;
        public SheetSelectionWindow(LayoutsToDWGViewModel vm)
        {
            InitializeComponent();
            _vm = vm;
            DataContext = vm;
        }

        public bool IsAllSelected
        {
            get => _vm.AllSheets.All(s => s.IsSelected);
            set
            {
                foreach (var s in _vm.AllSheets)
                    s.IsSelected = value;
            }
        }

        public System.Collections.Generic.List<Autodesk.Revit.DB.ViewSheet> SelectedSheets { get; private set; }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            SelectedSheets = _vm.AllSheets
                .Where(s => s.IsSelected)
                .Select(s => s.Sheet)
                .ToList();
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
