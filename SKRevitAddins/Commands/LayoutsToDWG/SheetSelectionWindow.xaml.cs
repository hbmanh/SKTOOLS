using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace LayoutsToDWG.Commands
{
    public partial class SheetSelectionWindow : Window
    {
        private readonly ViewModels _vm;

        public List<ViewSheet> SelectedSheets { get; private set; } = new();

        public SheetSelectionWindow(ViewModels vm)
        {
            InitializeComponent();
            _vm = vm;
            DataContext = _vm;
            _vm.LoadSheetSets();
            _vm.LoadSheets();
        }

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

        public bool IsAllSelected
        {
            get => _vm.AllSheets.All(s => s.IsSelected);
            set
            {
                foreach (var item in _vm.AllSheets)
                    item.IsSelected = value;
            }
        }
    }
}
