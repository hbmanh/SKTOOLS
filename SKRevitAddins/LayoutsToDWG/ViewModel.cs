using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Microsoft.Win32;
using SKRevitAddins.Utils;

namespace SKRevitAddins.LayoutsToDWG.ViewModel
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    public class SheetItem : ViewModelBase
    {
        public SheetItem(ViewSheet sheet, string num, string name)
        {
            Sheet = sheet;
            SheetNumber = num;
            SheetName = name;
        }

        public ViewSheet Sheet { get; }
        public string SheetNumber { get; }
        public string SheetName { get; }

        bool _isSel = true;
        public bool IsSelected { get => _isSel; set { _isSel = value; OnPropertyChanged(); } }

        int _order;
        public int Order { get => _order; set { _order = value; OnPropertyChanged(); } }
    }

    public class LayoutsToDWGViewModel : ViewModelBase
    {
        readonly UIDocument _uiDoc;
        Document Doc => _uiDoc.Document;

        readonly ExportSheetsHandler _handler = new();
        readonly ExternalEvent _evt;

        public LayoutsToDWGViewModel(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;
            _evt = ExternalEvent.Create(_handler);

            LoadExportSetups();
            LoadSheetSets();
            LoadTitleblockParams();

            BrowseFolderCmd = new RelayCommand(_ => BrowseFolder());
            ExportLayerCmd = new RelayCommand(_ => ExportLayerMapping());
            StartExportSheetsCmd = new RelayCommand(_ => StartExport());
        }

        //──────────────── Collections ──────────────────────────
        public ObservableCollection<string> ExportSetups { get; } = new();
        public ObservableCollection<string> SheetSets { get; } = new();
        public ObservableCollection<SheetItem> SheetItems { get; } = new();
        public ObservableCollection<string> TitleblockParams { get; } = new();

        //──────────────── Bindable props ───────────────────────
        public string SelectedExportSetup { get; set; }

        string _exportPath;
        public string ExportPath { get => _exportPath; set { _exportPath = value; OnPropertyChanged(); } }

        string _selSet;
        public string SelectedSheetSet
        {
            get => _selSet;
            set { _selSet = value; OnPropertyChanged(); LoadSheetsFromSet(); }
        }

        public string Prefix { get; set; }
        public string SheetNumberParam { get; set; }
        public string SheetNameParam { get; set; }
        public bool OpenFolderAfterExport { get; set; } = true;

        bool _isBusy;
        public bool IsBusy { get => _isBusy; set { _isBusy = value; OnPropertyChanged(); } }

        //──────────────── Commands ─────────────────────────────
        public ICommand BrowseFolderCmd { get; }
        public ICommand ExportLayerCmd { get; }
        public ICommand StartExportSheetsCmd { get; }

        //──────────────── Browse folder ───────────────────────
        void BrowseFolder()
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                ExportPath = dlg.SelectedPath;
        }

        //──────────────── Export layer mapping ────────────────
        void ExportLayerMapping()
        {
            if (string.IsNullOrWhiteSpace(SelectedExportSetup))
            {
                TaskDialog.Show("Export Layer Settings", "Vui lòng chọn Export Setup.");
                return;
            }

            var dlg = new SaveFileDialog
            {
                Title = "Save Layer Mapping",
                Filter = "Text File (*.txt)|*.txt",
                FileName = $"{SelectedExportSetup}.txt"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                LayerExportHelper.WriteLayerMapping(Doc, SelectedExportSetup, dlg.FileName);
                TaskDialog.Show("Completed", $"Đã xuất layer mapping:\n{dlg.FileName}");
            }
            catch (Exception ex) { TaskDialog.Show("Error", ex.Message); }
        }

        //──────────────── Start export DWG ─────────────────────
        void StartExport()
        {
            var sel = SheetItems.Where(x => x.IsSelected).ToList();
            if (!sel.Any()) { Msg("No sheet selected."); return; }
            if (!Directory.Exists(ExportPath)) { Msg("Path invalid."); return; }

            try
            {
                var dwgOpt = new FilteredElementCollector(Doc)
                                .OfClass(typeof(ExportDWGSettings))
                                .Cast<ExportDWGSettings>()
                                .First(x => x.Name == SelectedExportSetup)
                                .GetDWGExportOptions();

                _handler.ViewIds = sel.Select(x => x.Sheet.Id).ToList();
                _handler.TargetPath = ExportPath;
                _handler.Options = dwgOpt;
                _handler.FilePattern = $"{Prefix}-{{num}}_{{name}}";
                _handler.OpenFolder = OpenFolderAfterExport;
                _handler.BusySetter = v => IsBusy = v;

                _evt.Raise();   // chạy trong API context
            }
            catch (Exception ex) { Msg(ex.Message); }
        }

        //──────────────── Helper load data ─────────────────────
        void LoadExportSetups()
        {
            ExportSetups.Clear();
            foreach (var n in new FilteredElementCollector(Doc)
                                  .OfClass(typeof(ExportDWGSettings))
                                  .Cast<ExportDWGSettings>()
                                  .Select(x => x.Name)
                                  .OrderBy(n => n))
                ExportSetups.Add(n);
            SelectedExportSetup = ExportSetups.FirstOrDefault();
        }

        void LoadSheetSets()
        {
            SheetSets.Clear();
            foreach (var n in new FilteredElementCollector(Doc)
                                  .OfClass(typeof(ViewSheetSet))
                                  .Cast<ViewSheetSet>()
                                  .OrderBy(s => s.Name)
                                  .Select(s => s.Name))
                SheetSets.Add(n);
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
            int order = 1;
            foreach (var vs in set.Views.OfType<ViewSheet>())
                SheetItems.Add(new SheetItem(vs, vs.SheetNumber, vs.Name) { Order = order++ });
        }

        void LoadTitleblockParams()
        {
            var tb = new FilteredElementCollector(Doc)
                       .OfClass(typeof(FamilyInstance))
                       .Cast<FamilyInstance>()
                       .FirstOrDefault(fi => fi.Category?.Id.IntegerValue ==
                                             (int)BuiltInCategory.OST_TitleBlocks);
            if (tb == null) return;

            foreach (Parameter p in tb.Parameters)
                TitleblockParams.Add(p.Definition.Name);

            SheetNumberParam = TitleblockParams.FirstOrDefault();
            SheetNameParam = TitleblockParams.Skip(1).FirstOrDefault();
        }

        static void Msg(string t) => System.Windows.MessageBox.Show(t);
    }
}
