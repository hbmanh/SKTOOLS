using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using SKRevitAddins.Utils;

namespace SKRevitAddins.LayoutsToDWG.ViewModel
{
    #region ───────── Helpers ──────────────────────────────────
    public class InvertBoolConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object v, Type t, object p, System.Globalization.CultureInfo c)
            => !(bool)v;
        public object ConvertBack(object v, Type t, object p, System.Globalization.CultureInfo c)
            => !(bool)v;
    }

    public class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    public class SheetItem : ViewModelBase
    {
        public SheetItem(ViewSheet sheet) => Sheet = sheet;
        public ViewSheet Sheet { get; }
        public string SheetNumber => Sheet.SheetNumber;
        public string SheetName => Sheet.Name;

        bool _sel = true;
        public bool IsSelected { get => _sel; set { _sel = value; OnPropertyChanged(); } }
    }

    internal static class SettingsHelper
    {
        const string FOLDER = "SheetsToDWG";
        const string FILE = "settings.json";

        static string FullPath =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                         "Shinken", FOLDER, FILE);

        internal class ExportSettings
        {
            public string Prefix { get; set; }
            public string SheetNumberParam { get; set; }
            public string SheetNameParam { get; set; }
            public string ExportPath { get; set; }
        }

        public static ExportSettings Load()
        {
            try
            {
                if (File.Exists(FullPath))
                    return JsonSerializer.Deserialize<ExportSettings>(File.ReadAllText(FullPath))
                           ?? new ExportSettings();
            }
            catch { }
            return new ExportSettings();
        }

        public static void Save(ExportSettings s)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FullPath)!);
            File.WriteAllText(FullPath, JsonSerializer.Serialize(s,
                new JsonSerializerOptions { WriteIndented = true }));
        }
    }
    #endregion

    public class LayoutsToDWGViewModel : ViewModelBase
    {
        readonly UIDocument _uiDoc;
        readonly ExportSheetsHandler _handler = new();
        readonly ExternalEvent _evt;

        Document Doc => _uiDoc.Document;

        public LayoutsToDWGViewModel(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;
            _evt = ExternalEvent.Create(_handler);

            LoadExportSetups();
            LoadSheetSets();
            LoadTitleblockParams();
            LoadSaved();

            BrowseFolderCmd = new RelayCommand(_ => BrowseFolder(), _ => !IsBusy);
            ExportLayerCmd = new RelayCommand(_ => ExportLayerMapping(), _ => !IsBusy);
            StartExportSheetsCmd = new RelayCommand(_ => StartExport(), _ => !IsBusy);
        }

        #region ─── Collections & bindable ─────────────────────
        public ObservableCollection<string> ExportSetups { get; } = new();
        public ObservableCollection<string> SheetSets { get; } = new();
        public ObservableCollection<SheetItem> SheetItems { get; } = new();
        public ObservableCollection<string> TitleblockParams { get; } = new();

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

        double _progress;
        public double Progress { get => _progress; set { _progress = value; OnPropertyChanged(); } }

        bool _busy;
        public bool IsBusy
        {
            get => _busy;
            set { _busy = value; OnPropertyChanged(); CommandManager.InvalidateRequerySuggested(); }
        }
        #endregion

        #region ─── Commands ───────────────────────────────────
        public ICommand BrowseFolderCmd { get; }
        public ICommand ExportLayerCmd { get; }
        public ICommand StartExportSheetsCmd { get; }
        #endregion

        #region ─── Core: StartExport ─────────────────────────
        void StartExport()
        {
            if (IsBusy) return;

            var sel = SheetItems.Where(s => s.IsSelected).ToList();
            if (!sel.Any()) { Msg("No sheet selected."); return; }
            if (!Directory.Exists(ExportPath)) { Msg("Path invalid."); return; }

            // duplicate check
            if (!DupCheck(sel, $"{Prefix}-{{num}}_{{name}}", ExportPath, out string warn))
            { Msg(warn); return; }

            try
            {
                var dwgSet = new FilteredElementCollector(Doc)
                                .OfClass(typeof(ExportDWGSettings))
                                .Cast<ExportDWGSettings>()
                                .First(x => x.Name == SelectedExportSetup);
                var opt = dwgSet.GetDWGExportOptions();
                opt.MergedViews = false;          // always 1 sheet = 1 DWG

                _handler.ViewIds = sel.Select(s => s.Sheet.Id).ToList();
                _handler.TargetPath = ExportPath;
                _handler.Options = opt;
                _handler.FilePattern = $"{Prefix}-{{num}}_{{name}}";
                _handler.OpenFolder = OpenFolderAfterExport;
                _handler.BusySetter = v => IsBusy = v;
                _handler.ProgressReport = (c, t) => Progress = (double)c / t;

                Progress = 0;
                IsBusy = true;
                _evt.Raise();
            }
            catch (Exception ex) { Msg(ex.Message); }
        }
        #endregion

        #region ─── Helpers ───────────────────────────────────
        static bool DupCheck(IEnumerable<SheetItem> sheets, string pattern, string dir, out string msg)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in sheets)
            {
                string fname = LayerExportHelper.Sanitize(
                                   pattern.Replace("{num}", s.SheetNumber)
                                          .Replace("{name}", s.SheetName)) + ".dwg";

                if (!names.Add(fname)) { msg = $"Tên lặp: {fname}"; return false; }
                if (File.Exists(Path.Combine(dir, fname))) { msg = $"File tồn tại: {fname}"; return false; }
            }
            msg = null; return true;
        }

        void BrowseFolder()
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                ExportPath = dlg.SelectedPath;
        }

        void ExportLayerMapping()
        {
            if (string.IsNullOrEmpty(SelectedExportSetup))
            { TaskDialog.Show("Export Layer Settings", "Vui lòng chọn Export Setup."); return; }

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

        void LoadExportSetups()
        {
            foreach (var n in new FilteredElementCollector(Doc)
                              .OfClass(typeof(ExportDWGSettings))
                              .Cast<ExportDWGSettings>().Select(x => x.Name).OrderBy(n => n))
                ExportSetups.Add(n);
            SelectedExportSetup = ExportSetups.FirstOrDefault();
        }
        void LoadSheetSets()
        {
            foreach (var n in new FilteredElementCollector(Doc)
                              .OfClass(typeof(ViewSheetSet))
                              .Cast<ViewSheetSet>().Select(s => s.Name).OrderBy(n => n))
                SheetSets.Add(n);
            SelectedSheetSet = SheetSets.FirstOrDefault();
        }
        void LoadSheetsFromSet()
        {
            var set = new FilteredElementCollector(Doc)
                        .OfClass(typeof(ViewSheetSet))
                        .Cast<ViewSheetSet>()
                        .FirstOrDefault(s => s.Name == SelectedSheetSet);
            if (set == null) return;

            SheetItems.Clear();
            foreach (var vs in set.Views.OfType<ViewSheet>())
                SheetItems.Add(new SheetItem(vs));
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

        void LoadSaved()
        {
            var s = SettingsHelper.Load();
            if (s == null) return;
            if (s.Prefix != null) Prefix = s.Prefix;
            if (s.SheetNumberParam != null) SheetNumberParam = s.SheetNumberParam;
            if (s.SheetNameParam != null) SheetNameParam = s.SheetNameParam;
            if (s.ExportPath != null) ExportPath = s.ExportPath;
        }

        public void SaveCurrent()
        {
            SettingsHelper.Save(new SettingsHelper.ExportSettings
            {
                Prefix = Prefix,
                SheetNumberParam = SheetNumberParam,
                SheetNameParam = SheetNameParam,
                ExportPath = ExportPath
            });
        }
        #endregion

        static void Msg(string t) => MessageBox.Show(t);
    }
}
