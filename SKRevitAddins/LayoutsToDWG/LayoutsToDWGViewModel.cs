using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;          // RelayCommand
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Input;

namespace SKRevitAddins.LayoutsToDWG
{
    public sealed class LayoutsToDWGViewModel : INotifyPropertyChanged
    {
        // -------- ctor & fields --------
        private readonly UIDocument _uiDoc;
        private readonly ExternalEvent _externalEvent;
        private readonly LayoutsToDWGRequestHandler _handler;

        public LayoutsToDWGViewModel(UIDocument uiDoc,
                                     ExternalEvent externalEvent,
                                     LayoutsToDWGRequestHandler handler)
        {
            _uiDoc = uiDoc;
            _externalEvent = externalEvent;
            _handler = handler;

            LoadExportSetups();
            LoadTitleblockParams();

            BrowseFolderCmd = new RelayCommand(_ => OnBrowseFolder());
            ExportLayerSettingsCmd = new RelayCommand(_ => OnExportLayerSettings());
        }

        // -------- EXPORT OPTION --------
        public ObservableCollection<string> ExportSetups { get; } = new();

        private string? _selectedExportSetup;
        public string? SelectedExportSetup
        {
            get => _selectedExportSetup;
            set => Set(ref _selectedExportSetup, value);
        }

        private string _exportFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public string ExportFolder
        {
            get => _exportFolder;
            set => Set(ref _exportFolder, value);
        }

        // -------- FILE NAME STRUCTURE ---
        private string _prefix = "";
        public string Prefix
        {
            get => _prefix;
            set => Set(ref _prefix, value);
        }

        public ObservableCollection<string> TitleblockParams { get; } = new();

        private string? _sheetNumberParam;
        public string? SheetNumberParam
        {
            get => _sheetNumberParam;
            set => Set(ref _sheetNumberParam, value);
        }

        private string? _sheetNameParam;
        public string? SheetNameParam
        {
            get => _sheetNameParam;
            set => Set(ref _sheetNameParam, value);
        }

        // -------- OPEN FOLDER ----------
        private bool _openAfterExport = true;
        public bool OpenAfterExport
        {
            get => _openAfterExport;
            set => Set(ref _openAfterExport, value);
        }

        // -------- COMMANDS -------------
        public ICommand BrowseFolderCmd { get; }
        public ICommand ExportLayerSettingsCmd { get; }

        private void OnBrowseFolder()
        {
            using var dlg = new FolderBrowserDialog
            {
                Description = "Chọn thư mục lưu DWG",
                SelectedPath = ExportFolder,
                ShowNewFolderButton = true
            };
            if (dlg.ShowDialog() == DialogResult.OK)
                ExportFolder = dlg.SelectedPath;
        }

        private string? _layerTxtPath;               // đường dẫn người dùng chọn

        private void OnExportLayerSettings()
        {
            using var dlg = new SaveFileDialog
            {
                Title = "Lưu Layer Mapping",
                Filter = "Revit layer map (*.txt)|*.txt|CSV (*.csv)|*.csv",
                FileName = "exportlayers-dwg-AIA.txt",
                DefaultExt = "txt"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            _layerTxtPath = dlg.FileName;
            _externalEvent.Raise();                  // Handler biết cần export layer
        }

        // -------- INIT DATA ------------
        private void LoadExportSetups()
        {
            var names = ExportUtils.GetExportDWGSettings(_uiDoc.Document)
                        .Select(s => s.Name)
                        .OrderBy(s => s)
                        .ToList();

            ExportSetups.Clear();
            foreach (var n in names) ExportSetups.Add(n);

            SelectedExportSetup = names.FirstOrDefault();
        }

        private void LoadTitleblockParams()
        {
            var titleblocks = new FilteredElementCollector(_uiDoc.Document)
                              .OfClass(typeof(FamilyInstance))
                              .OfCategory(BuiltInCategory.OST_TitleBlocks)
                              .Cast<FamilyInstance>();

            var names = titleblocks
                        .SelectMany(tb => tb.Parameters.Cast<Parameter>())
                        .Select(p => p.Definition.Name)
                        .Distinct()
                        .OrderBy(n => n)
                        .ToList();

            TitleblockParams.Clear();
            foreach (var n in names) TitleblockParams.Add(n);

            SheetNumberParam = names.FirstOrDefault(
                n => n.IndexOf("Number", StringComparison.OrdinalIgnoreCase) >= 0);

            SheetNameParam = names.FirstOrDefault(
                n => n.IndexOf("Title", StringComparison.OrdinalIgnoreCase) >= 0)
                ?? names.FirstOrDefault(
                n => n.IndexOf("Name", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        // -------- INotify --------------
        public event PropertyChangedEventHandler? PropertyChanged;
        private bool Set<T>(ref T field, T value,
                            [CallerMemberName] string? prop = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
            return true;
        }

        // -------- API cho Handler ------
        private IList<ElementId> _selectedSheetIds = new List<ElementId>();
        public void SetSelectedSheets(IEnumerable<ElementId> ids) =>
            _selectedSheetIds = ids.ToList();

        internal LayoutsToDWGRequest CreateRequest() => new()
        {
            SheetIds = _selectedSheetIds,        // có thể rỗng
            ExportSetupName = SelectedExportSetup,
            ExportFolder = ExportFolder,

            Prefix = Prefix,
            SheetNumberParam = SheetNumberParam,
            SheetNameParam = SheetNameParam,

            LayerTxtPath = _layerTxtPath,
            OpenAfterExport = OpenAfterExport
        };

        internal void ResetLayerTxt() => _layerTxtPath = null;  // Handler gọi
    }
}
