using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using SKRevitAddins.Utils;        // RelayCommand

namespace SKRevitAddins.LayoutsToDWG.ViewModel
{
    //───────────────────────── 1.  Base  ─────────────────────────
    public class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    //───────────────────────── 2.  Item cho DataGrid  ────────────
    public class SheetItem : ViewModelBase
    {
        public ViewSheet Sheet { get; set; }
        public ElementId SheetId => Sheet?.Id;
        public string SheetNumber { get; set; }
        public string SheetName { get; set; }

        bool _isSel;
        public bool IsSelected { get => _isSel; set { _isSel = value; OnPropertyChanged(); } }

        int _order;
        public int Order { get => _order; set { _order = value; OnPropertyChanged(); } }
    }

    //───────────────────────── 3.  MAIN ViewModel  ───────────────
    public class LayoutsToDWGViewModel : ViewModelBase
    {
        readonly UIDocument _uiDoc;
        public Document Doc => _uiDoc.Document;

        //── ctor ────────────────────────────────────────────────
        public LayoutsToDWGViewModel(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;

            SheetItems = new ObservableCollection<SheetItem>();
            SheetSets = new ObservableCollection<string>();
            ExportSetups = new ObservableCollection<string>();

            // DWG Export setups
            foreach (var s in new FilteredElementCollector(Doc)
                                .OfClass(typeof(ExportDWGSettings))
                                .Cast<ExportDWGSettings>()
                                .Select(x => x.Name)
                                .OrderBy(n => n))
                ExportSetups.Add(s);

            SelectedExportSetup = ExportSetups.FirstOrDefault();
            OpenFolderAfterExport = true;

            LoadSheetSets();
            LoadTitleblockParams();          // populate TitleblockParams

            //──‑‑‑‑‑ Mặc định Middle/Ending ‑‑‑‑‑────────────────
            SheetNumberParam = TitleblockParams
                .FirstOrDefault(n => n.Equals("Sheet Number", StringComparison.OrdinalIgnoreCase))
                ?? TitleblockParams.FirstOrDefault();

            SheetNameParam = TitleblockParams
                .FirstOrDefault(n => n.Equals("Sheet Name", StringComparison.OrdinalIgnoreCase))
                ?? TitleblockParams.Skip(1).FirstOrDefault();

            ExportLayerSettingsCmd = new RelayCommand(_ => ExportLayerSettings());
        }

        //── Collections ─────────────────────────────────────────
        public ObservableCollection<SheetItem> SheetItems { get; }
        public ObservableCollection<string> SheetSets { get; }
        public ObservableCollection<string> ExportSetups { get; }
        public ObservableCollection<string> TitleblockParams { get; } = new();

        //── Sheet‑set lựa chọn ──────────────────────────────────
        string _selSet;
        public string SelectedSheetSet
        {
            get => _selSet;
            set { _selSet = value; OnPropertyChanged(); LoadSheetsFromSet(); }
        }

        //── DWG export setup ────────────────────────────────────
        public string SelectedExportSetup { get; set; }

        //── Export path ─────────────────────────────────────────
        string _exportPath;
        public string ExportPath
        {
            get => _exportPath;
            set { _exportPath = value; OnPropertyChanged(); }
        }

        //── Mở Explorer sau export ──────────────────────────────
        bool _openFolder;
        public bool OpenFolderAfterExport
        {
            get => _openFolder;
            set { _openFolder = value; OnPropertyChanged(); }
        }

        //──────── FILE NAME STRUCTURE ───────────────────────────
        public string Prefix { get; set; }

        string _sheetNumParam;
        public string SheetNumberParam
        {
            get => _sheetNumParam;
            set { _sheetNumParam = value; OnPropertyChanged(); }
        }

        string _sheetNameParam;
        public string SheetNameParam
        {
            get => _sheetNameParam;
            set { _sheetNameParam = value; OnPropertyChanged(); }
        }

        //──────── ICommand: Export layer settings ───────────────
        public ICommand ExportLayerSettingsCmd { get; }

        void ExportLayerSettings()
        {
            TaskDialog.Show("Export Layer Settings",
                            "Chức năng đang được phát triển.");
        }

        //───────────────────────── 4.  Helper methods ───────────
        void LoadSheetSets()
        {
            SheetSets.Clear();
            foreach (var s in new FilteredElementCollector(Doc)
                                  .OfClass(typeof(ViewSheetSet))
                                  .Cast<ViewSheetSet>()
                                  .OrderBy(s => s.Name))
                SheetSets.Add(s.Name);

            SelectedSheetSet = SheetSets.FirstOrDefault();
        }

        void LoadSheetsFromSet()
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
                    Order = idx++,
                    IsSelected = true
                });
            }
        }

        void LoadTitleblockParams()
        {
            // Lấy 1 titleblock bất kỳ để liệt kê Parameter
            var tb = new FilteredElementCollector(Doc)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .FirstOrDefault(fi => fi.Category?.Id.IntegerValue ==
                                              (int)BuiltInCategory.OST_TitleBlocks);
            if (tb == null) return;

            foreach (Parameter p in tb.Parameters)
            {
                string name = p.Definition.Name;
                if (!TitleblockParams.Contains(name))
                    TitleblockParams.Add(name);
            }
        }
    }
}
