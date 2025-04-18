using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Input;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    public class LayoutsToDWGViewModel : INotifyPropertyChanged
    {
        readonly UIApplication _app;
        readonly Document _doc;

        private readonly List<SheetSelectionItemLight> _preloadedItems = new();

        private string _selectedExportSetup;
        private string _numberLabel;
        private string _nameLabel;
        private string _scaleLabel;
        private string _selectedSheetSet;
        private bool _isMergeMode;
        private bool _openAfterExport;
        private string _exportFolder;
        private string _mergedFilename;

        // AutoCAD merge flags
        public bool MergeLayers { get; set; } = true;
        public bool RasterToOle { get; set; } = false;
        public bool IsPremium { get; set; } = false;
        public string CadPath { get; set; } = "";

        public LayoutsToDWGViewModel(UIApplication app)
        {
            _app = app;
            _doc = app.ActiveUIDocument.Document;

            SelectionItems = new ObservableCollection<SheetSelectionItemLight>();
            ExportSetups = new ObservableCollection<string>();
            TitleblockParams = new ObservableCollection<string>();
            SheetSets = new ObservableCollection<string>();

            BrowseFolderCmd = new RelayCommand(_ => BrowseFolder());

            // Defaults
            IsMergeMode = false;
            OpenAfterExport = true;
            MergedFilename = "MergedSheets.dwg";
            ExportFolder = string.Empty;

            LoadExportSetups();
            LoadTitleblockParams();
            LoadSheetSets();
            PreloadSheets();
        }

        public ObservableCollection<SheetSelectionItemLight> SelectionItems { get; }
        public ObservableCollection<string> ExportSetups { get; }
        public ObservableCollection<string> TitleblockParams { get; }
        public ObservableCollection<string> SheetSets { get; }
        public ICommand BrowseFolderCmd { get; }

        public string SelectedExportSetup
        {
            get => _selectedExportSetup;
            set => Set(ref _selectedExportSetup, value);
        }

        public void LoadExportSetups()
        {
            ExportSetups.Clear();
            foreach (ExportDWGSettings s in new FilteredElementCollector(_doc)
                     .OfClass(typeof(ExportDWGSettings)))
                ExportSetups.Add(s.Name);
            SelectedExportSetup = ExportSetups.FirstOrDefault();
        }

        public string NumberLabel
        {
            get => _numberLabel;
            set => Set(ref _numberLabel, value);
        }
        public string NameLabel
        {
            get => _nameLabel;
            set => Set(ref _nameLabel, value);
        }
        public string ScaleLabel
        {
            get => _scaleLabel;
            set => Set(ref _scaleLabel, value);
        }

        public void LoadTitleblockParams()
        {
            TitleblockParams.Clear();
            var titles = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType();

            var names = new HashSet<string>();
            foreach (Element el in titles)
                foreach (Parameter p in el.Parameters.Cast<Parameter>())
                    if (p.StorageType == StorageType.String && p.HasValue)
                        names.Add(p.Definition.Name);

            foreach (var n in names.OrderBy(n => n))
                TitleblockParams.Add(n);

            NumberLabel = TitleblockParams.Contains("Sheet Number") ? "Sheet Number" : TitleblockParams.FirstOrDefault();
            NameLabel = TitleblockParams.Contains("Sheet Name") ? "Sheet Name" : TitleblockParams.FirstOrDefault();
            ScaleLabel = TitleblockParams.Contains("Scale") ? "Scale" : TitleblockParams.FirstOrDefault();
        }

        public string SelectedSheetSet
        {
            get => _selectedSheetSet;
            set
            {
                if (Set(ref _selectedSheetSet, value))
                    RefreshSelectionItems();
            }
        }
        public void LoadSheetSets()
        {
            SheetSets.Clear();
            SheetSets.Add("- By Default -");
            foreach (ViewSheetSet vs in new FilteredElementCollector(_doc)
                     .OfClass(typeof(ViewSheetSet)))
                SheetSets.Add(vs.Name);
            SelectedSheetSet = SheetSets.FirstOrDefault();
        }

        private void PreloadSheets()
        {
            _preloadedItems.Clear();
            var sheets = new FilteredElementCollector(_doc)
                         .OfClass(typeof(ViewSheet))
                         .Cast<ViewSheet>();
            foreach (var s in sheets)
                _preloadedItems.Add(new SheetSelectionItemLight
                {
                    Id = s.Id,
                    SheetName = $"{s.SheetNumber} - {s.Name}",
                    IsSelected = false
                });
            RefreshSelectionItems();
        }

        private void RefreshSelectionItems()
        {
            var filtered = _preloadedItems.AsEnumerable();
            if (SelectedSheetSet != "- By Default -")
            {
                var set = new FilteredElementCollector(_doc)
                          .OfClass(typeof(ViewSheetSet))
                          .Cast<ViewSheetSet>()
                          .FirstOrDefault(x => x.Name == SelectedSheetSet);
                if (set != null)
                {
                    var ids = new HashSet<ElementId>(
                        set.Views.Cast<Autodesk.Revit.DB.View>().Select(v => v.Id)
                    );
                    filtered = filtered.Where(i => ids.Contains(i.Id));
                }
            }

            SelectionItems.Clear();
            foreach (var item in filtered)
            {
                item.IsSelected = SelectedSheetSet != "- By Default -";
                SelectionItems.Add(item);
            }
        }

        public List<ViewSheet> SelectedSheets { get; private set; }
        public void PrepareSelectedSheets()
        {
            SelectedSheets = SelectionItems
                .Where(i => i.IsSelected)
                .Select(i => _doc.GetElement(i.Id) as ViewSheet)
                .Where(v => v != null)
                .ToList();
        }

        public bool IsMergeMode
        {
            get => _isMergeMode;
            set => Set(ref _isMergeMode, value);
        }
        public bool IsPerSheetMode => !_isMergeMode;

        public bool OpenAfterExport
        {
            get => _openAfterExport;
            set => Set(ref _openAfterExport, value);
        }

        public string ExportFolder
        {
            get => _exportFolder;
            set => Set(ref _exportFolder, value);
        }
        private void BrowseFolder()
        {
            using var dlg = new FolderBrowserDialog { Description = "Select export folder" };
            if (dlg.ShowDialog() == DialogResult.OK)
                ExportFolder = dlg.SelectedPath;
        }

        public string MergedFilename
        {
            get => _mergedFilename;
            set => Set(ref _mergedFilename, value);
        }

        public event PropertyChangedEventHandler PropertyChanged;
        void OnPropertyChanged([CallerMemberName] string prop = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));

        bool Set<T>(ref T field, T value, [CallerMemberName] string prop = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(prop);
            return true;
        }
    }
}
