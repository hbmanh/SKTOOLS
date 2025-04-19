// SKRevitAddins.Commands.LayoutsToDWG.ViewModel.cs
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Data;

namespace SKRevitAddins.Commands.LayoutsToDWG.ViewModel
{
    // ───────────── Base + INotifyPropertyChanged ─────────────
    public class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
    }

    // ───────────── Converter ─────────────
    public class InverseBoolConverter : IValueConverter
    {
        public object Convert(object value, System.Type targetType, object parameter, CultureInfo culture)
            => value is bool b ? !b : false;

        public object ConvertBack(object value, System.Type targetType, object parameter, CultureInfo culture)
            => value is bool b ? !b : false;
    }

    // ───────────── SheetItem (row) ─────────────
    public class SheetItem : ViewModelBase
    {
        public ViewSheet Sheet { get; set; }
        public ElementId SheetId => Sheet?.Id;
        public string SheetNumber { get; set; }
        public string SheetName { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        private int _order;
        public int Order
        {
            get => _order;
            set { _order = value; OnPropertyChanged(); }
        }
    }

    // ───────────── Main ViewModel ─────────────
    public class LayoutsToDWGViewModel : ViewModelBase
    {
        readonly UIDocument _uiDoc;
        public Document Doc => _uiDoc.Document;

        public LayoutsToDWGViewModel(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;
            SheetItems = new ObservableCollection<SheetItem>();
            ExportSetups = new ObservableCollection<string>(
                new FilteredElementCollector(Doc)
                    .OfClass(typeof(ExportDWGSettings))
                    .Cast<ExportDWGSettings>()
                    .Select(x => x.Name)
                    .OrderBy(n => n)
            );
            SelectedExportSetup = ExportSetups.FirstOrDefault();
            OpenFolderAfterExport = true;
            MergeSheets = true;

            LoadSheetSets();
            UpdateMergeFilename();
        }

        // Collections
        public ObservableCollection<SheetItem> SheetItems { get; }
        public ObservableCollection<string> SheetSets { get; } = new();
        public ObservableCollection<string> ExportSetups { get; }

        // Properties
        private string _selectedSheetSet;
        public string SelectedSheetSet
        {
            get => _selectedSheetSet;
            set
            {
                _selectedSheetSet = value;
                OnPropertyChanged();
                LoadSheetsFromSet();
                UpdateMergeFilename();
            }
        }

        public string SelectedExportSetup { get; set; }

        private string _exportPath;
        public string ExportPath
        {
            get => _exportPath;
            set { _exportPath = value; OnPropertyChanged(); }
        }

        private string _mergeFilename;
        public string MergeFilename
        {
            get => _mergeFilename;
            set { _mergeFilename = value; OnPropertyChanged(); }
        }

        private bool _mergeSheets;
        public bool MergeSheets
        {
            get => _mergeSheets;
            set
            {
                _mergeSheets = value;
                OnPropertyChanged();
                if (_mergeSheets) UpdateMergeFilename();
            }
        }

        private bool _orientationHorizontal;
        public bool OrientationHorizontal
        {
            get => _orientationHorizontal;
            set { _orientationHorizontal = value; OnPropertyChanged(); }
        }

        private bool _openFolderAfterExport;
        public bool OpenFolderAfterExport
        {
            get => _openFolderAfterExport;
            set { _openFolderAfterExport = value; OnPropertyChanged(); }
        }

        // Load sheet‐sets & sheets
        private void LoadSheetSets()
        {
            var sets = new FilteredElementCollector(Doc)
                .OfClass(typeof(ViewSheetSet))
                .Cast<ViewSheetSet>()
                .OrderBy(s => s.Name);

            SheetSets.Clear();
            foreach (var s in sets) SheetSets.Add(s.Name);
            SelectedSheetSet = SheetSets.FirstOrDefault();
        }

        private void LoadSheetsFromSet()
        {
            if (string.IsNullOrEmpty(SelectedSheetSet)) return;
            var set = new FilteredElementCollector(Doc)
                .OfClass(typeof(ViewSheetSet))
                .Cast<ViewSheetSet>()
                .FirstOrDefault(s => s.Name == SelectedSheetSet);
            if (set == null) return;

            SheetItems.Clear();
            int idx = 1;
            foreach (var vs in set.Views.Cast<ViewSheet>().OrderBy(v => v.SheetNumber))
            {
                SheetItems.Add(new SheetItem
                {
                    Sheet = vs,
                    SheetNumber = vs.SheetNumber,
                    SheetName = vs.Name,
                    IsSelected = true,
                    Order = idx++
                });
            }
        }

        // Merge‐filename helper
        public void UpdateMergeFilename()
        {
            if (!MergeSheets) return;
            var nums = SheetItems
                .Where(si => si.IsSelected)
                .Select(si => si.SheetNumber)
                .OrderBy(n => n)
                .ToList();
            if (nums.Count > 0)
                MergeFilename = nums.Count == 1
                    ? nums.First()
                    : $"{nums.First()}-{nums.Last()}";
        }
    }
}
