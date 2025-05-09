using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using SKRevitAddins.LayoutsToDWG;
using SKRevitAddins.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace SKRevitAddins.LayoutsToDWG.ViewModel
{
    /// <summary>
    /// Đảo ngược bool cho binding trong XAML
    /// </summary>
    public class InvertBoolConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object v, Type t, object p, System.Globalization.CultureInfo c)
            => !(bool)v;
        public object ConvertBack(object v, Type t, object p, System.Globalization.CultureInfo c)
            => !(bool)v;
    }

    /// <summary>
    /// Base cho INotifyPropertyChanged
    /// </summary>
    public class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    /// <summary>
    /// Item cho mỗi sheet trong DataGrid
    /// </summary>
    public class SheetItem : ViewModelBase
    {
        public SheetItem(ViewSheet s) => Sheet = s;
        public ViewSheet Sheet { get; }
        public string SheetNumber => Sheet.SheetNumber;
        public string SheetName => Sheet.Name;

        bool _sel = true;
        public bool IsSelected
        {
            get => _sel;
            set { _sel = value; OnPropertyChanged(); }
        }
    }

    /// <summary>
    /// Item cho cấu hình tên file
    /// </summary>
    public class FileNameItem : ViewModelBase
    {
        public FileNameItem(string param, string sep)
        {
            SelectedParam = param;
            Sep = sep;
        }

        public string SelectedParam { get => _param; set { _param = value; OnPropertyChanged(); } }
        public string Sep { get => _sep; set { _sep = value; OnPropertyChanged(); } }

        string _param, _sep;
    }

    /// <summary>
    /// So sánh "tự nhiên" cho SheetNumber
    /// </summary>
    class NaturalSheetComparer : IComparer<string>
    {
        static readonly Regex Rx = new(@"(\d+)|(\D+)", RegexOptions.Compiled);
        public int Compare(string a, string b)
        {
            if (a == null) return b == null ? 0 : -1;
            if (b == null) return 1;
            var ma = Rx.Matches(a);
            var mb = Rx.Matches(b);
            for (int i = 0; i < Math.Min(ma.Count, mb.Count); i++)
            {
                var xa = ma[i].Value;
                var xb = mb[i].Value;
                if (int.TryParse(xa, out int na) && int.TryParse(xb, out int nb))
                {
                    int c = na.CompareTo(nb);
                    if (c != 0) return c;
                }
                else
                {
                    int c = string.CompareOrdinal(xa, xb);
                    if (c != 0) return c;
                }
            }
            return ma.Count.CompareTo(mb.Count);
        }
    }

    /// <summary>
    /// ViewModel chính
    /// </summary>
    public class LayoutsToDWGViewModel : ViewModelBase
    {
        readonly UIDocument _uiDoc;
        readonly ExportSheetsHandler _handler;
        readonly ExternalEvent _evt;

        Document Doc => _uiDoc.Document;

        // --- Cho multi-set ---
        public ObservableCollection<string> SheetSets { get; } = new();
        public ObservableCollection<string> SelectedSheetSets { get; } = new();
        public ObservableCollection<SheetItem> SheetItems { get; } = new();

        // Lưu mapping sheetId → tên set
        readonly Dictionary<ElementId, string> _sheetSetMap = new();

        // --- Các collection & cache khác ---
        public ObservableCollection<string> ExportSetups { get; } = new();
        public ObservableCollection<FileNameItem> FileNameItems { get; } = new();
        public ObservableCollection<string> ParamOptions { get; } = new();

        public string SelectedExportSetup { get; set; }

        string _exportPath;
        public string ExportPath
        {
            get => _exportPath;
            set { _exportPath = value; OnPropertyChanged(); }
        }

        public bool OpenFolderAfterExport { get; set; } = true;

        bool _busy;
        public bool IsBusy
        {
            get => _busy;
            set { _busy = value; OnPropertyChanged(); CommandManager.InvalidateRequerySuggested(); }
        }

        double _progress;
        public double Progress
        {
            get => _progress;
            set { _progress = value; OnPropertyChanged(); }
        }

        // Commands
        public ICommand BrowseFolderCmd { get; }
        public ICommand ExportLayerCmd { get; }
        public ICommand StartExportSheetsCmd { get; }
        public ICommand CancelCmd { get; }
        public ICommand CancelExportCmd { get; }

        public Action RequestClose { get; set; }

        // --- Caches để tăng tốc lookup parameter & title-block ---
        readonly ILookup<ElementId, FamilyInstance> _titleBlocksBySheet;
        readonly Dictionary<(ElementId, string), string> _paramValueCache = new();

        public LayoutsToDWGViewModel(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;

            // 1-lần build lookup title-block theo sheet
            _titleBlocksBySheet = new FilteredElementCollector(Doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(tb => tb.OwnerViewId != ElementId.InvalidElementId)
                .ToLookup(tb => tb.OwnerViewId);

            _handler = new ExportSheetsHandler();
            _evt = ExternalEvent.Create(_handler);

            LoadExportSetups();
            LoadSheetSets();

            // Khi user chọn/deselect trong ListBox → rebuild sheets
            SelectedSheetSets.CollectionChanged += (_, __) => LoadSheetsFromSets();

            LoadParamOptions();
            InitFileNameItems();
            LoadSavedSettings();

            BrowseFolderCmd = new RelayCommand(_ => BrowseFolder(), _ => !IsBusy);
            ExportLayerCmd = new RelayCommand(_ => ExportLayerMapping(), _ => !IsBusy);
            StartExportSheetsCmd = new RelayCommand(_ => StartExport(), _ => !IsBusy);
            CancelExportCmd = new RelayCommand(_ => CancelExport(), _ => IsBusy);
            CancelCmd = new RelayCommand(_ => RequestClose?.Invoke());
        }

        #region Load danh sách và sheets

        void LoadExportSetups()
        {
            ExportSetups.Clear();
            foreach (var name in new FilteredElementCollector(Doc)
                       .OfClass(typeof(ExportDWGSettings))
                       .Cast<ExportDWGSettings>()
                       .Select(x => x.Name)
                       .OrderBy(n => n))
                ExportSetups.Add(name);

            SelectedExportSetup = ExportSetups.FirstOrDefault();
        }

        void LoadSheetSets()
        {
            SheetSets.Clear();
            foreach (var name in new FilteredElementCollector(Doc)
                       .OfClass(typeof(ViewSheetSet))
                       .Cast<ViewSheetSet>()
                       .Select(s => s.Name)
                       .OrderBy(n => n))
                SheetSets.Add(name);
        }

        void LoadSheetsFromSets()
        {
            SheetItems.Clear();
            _sheetSetMap.Clear();

            var cmp = new NaturalSheetComparer();
            foreach (var setName in SelectedSheetSets)
            {
                var set = new FilteredElementCollector(Doc)
                            .OfClass(typeof(ViewSheetSet))
                            .Cast<ViewSheetSet>()
                            .FirstOrDefault(s => s.Name == setName);
                if (set == null) continue;

                foreach (var vs in set.Views.OfType<ViewSheet>()
                            .OrderBy(v => v.SheetNumber, cmp))
                {
                    SheetItems.Add(new SheetItem(vs));
                    _sheetSetMap[vs.Id] = setName;
                }
            }
        }

        #endregion

        #region Build file name & lookup parameter

        string BuildFileName(ViewSheet sheet)
        {
            var parts = new List<string>();
            for (int i = 0; i < FileNameItems.Count; i++)
            {
                var item = FileNameItems[i];
                string value;

                if (item.SelectedParam == "Sheet Number")
                    value = sheet.SheetNumber;
                else if (item.SelectedParam == "Sheet Name")
                    value = sheet.LookupParameter("Sheet Name")?.AsString() ?? sheet.Name;
                else
                {
                    var p = sheet.LookupParameter(item.SelectedParam);
                    if (p != null && p.HasValue)
                        value = p.AsString();
                    else
                        value = GetParamValue(sheet, item.SelectedParam);
                }

                value = LayerExportHelper.Sanitize(value?.Trim() ?? "");
                if (string.IsNullOrWhiteSpace(value)) continue;

                parts.Add(value);

                // Lookahead: nếu còn phần tiếp theo có giá trị → thêm Sep
                bool hasNext = FileNameItems
                    .Skip(i + 1)
                    .Any(n => !string.IsNullOrWhiteSpace(
                        n.SelectedParam == "Sheet Number" ? sheet.SheetNumber
                      : n.SelectedParam == "Sheet Name" ? (sheet.LookupParameter("Sheet Name")?.AsString() ?? sheet.Name)
                      : GetParamValue(sheet, n.SelectedParam)));

                if (!string.IsNullOrWhiteSpace(item.Sep) && hasNext)
                    parts.Add(item.Sep);
            }

            var final = string.Concat(parts);

            // Trim separator dư cuối
            var lastSep = FileNameItems.LastOrDefault(f => !string.IsNullOrWhiteSpace(f.Sep))?.Sep;
            if (!string.IsNullOrWhiteSpace(lastSep) && final.EndsWith(lastSep))
                final = final.Substring(0, final.Length - lastSep.Length);

            return final;
        }

        string GetParamValue(ViewSheet sheet, string paramName)
        {
            var key = (sheet.Id, paramName);
            if (_paramValueCache.TryGetValue(key, out var cached))
                return cached;

            // 1) trực tiếp trên sheet
            var pSheet = sheet.LookupParameter(paramName);
            if (pSheet != null && pSheet.HasValue)
                return _paramValueCache[key] = pSheet.AsString();

            // 2) title-block đã cache
            if (_titleBlocksBySheet.Contains(sheet.Id))
            {
                var tbParam = _titleBlocksBySheet[sheet.Id]
                    .Select(tb => tb.LookupParameter(paramName))
                    .FirstOrDefault(x => x != null && x.HasValue);
                if (tbParam != null)
                    return _paramValueCache[key] = tbParam.AsString();
            }

            // 3) fallback scan project once
            var projParam = new FilteredElementCollector(Doc)
                .WhereElementIsNotElementType()
                .Select(e => e.LookupParameter(paramName))
                .FirstOrDefault(pp => pp != null && pp.HasValue);
            if (projParam != null)
                return _paramValueCache[key] = projParam.AsString();

            return _paramValueCache[key] = string.Empty;
        }

        #endregion

        #region Export / Cancel / Layer mapping

        void CancelExport()
        {
            _handler.IsCancelled = true;
            Msg("Cancelling export...");
        }

        void BrowseFolder()
        {
            using var dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
                ExportPath = dlg.SelectedPath;
        }

        void ExportLayerMapping()
        {
            if (string.IsNullOrWhiteSpace(SelectedExportSetup))
            {
                TaskDialog.Show("Layer", "Choose export setup.");
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "TXT|*.txt",
                FileName = $"{SelectedExportSetup}.txt"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                LayerExportHelper.WriteLayerMapping(Doc, SelectedExportSetup, dlg.FileName);
                TaskDialog.Show("Completed", dlg.FileName);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        void StartExport()
        {
            if (IsBusy) return;

            var sel = SheetItems.Where(s => s.IsSelected).ToList();
            if (!sel.Any()) { Msg("No sheet selected."); return; }
            if (!Directory.Exists(ExportPath)) { Msg("Path invalid."); return; }

            // Build map sheet → fileName
            var fileNames = sel.ToDictionary(
                s => s.Sheet.Id,
                s => BuildFileName(s.Sheet));

            if (!CheckDuplicates(fileNames, ExportPath, out string warn))
            {
                Msg(warn);
                return;
            }

            var setup = new FilteredElementCollector(Doc)
                          .OfClass(typeof(ExportDWGSettings))
                          .Cast<ExportDWGSettings>()
                          .FirstOrDefault(x => x.Name == SelectedExportSetup);
            if (setup == null) { Msg("Export setup not found."); return; }

            var opt = setup.GetDWGExportOptions();
            opt.MergedViews = true;

            // Build map sheet → subfolder (set name)
            var subFolders = new Dictionary<ElementId, string>();
            foreach (var id in fileNames.Keys)
                if (_sheetSetMap.TryGetValue(id, out var setName))
                    subFolders[id] = setName;

            // Gán vào handler
            _handler.ViewIds = fileNames.Keys.ToList();
            _handler.FileNames = fileNames;
            _handler.SubFolders = subFolders;
            _handler.TargetPath = ExportPath;
            _handler.Options = opt;
            _handler.OpenFolder = OpenFolderAfterExport;
            _handler.BusySetter = v => IsBusy = v;
            _handler.ProgressReporter = (c, t) => Progress = t == 0 ? 0 : (double)c / t;

            IsBusy = true;
            _evt.Raise();
            SaveCurrentSettings();
        }

        static bool CheckDuplicates(
            Dictionary<ElementId, string> fn,
            string dir,
            out string msg)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in fn.Values)
            {
                var f = name + ".dwg";
                if (!names.Add(f)) { msg = $"Duplicate: {f}"; return false; }
                if (File.Exists(Path.Combine(dir, f))) { msg = $"Exists:   {f}"; return false; }
            }
            msg = null; return true;
        }

        #endregion

        #region Settings Lưu/Load

        void LoadParamOptions()
        {
            ParamOptions.Clear();
            var names = new HashSet<string>();

            // Title-block
            foreach (var tb in new FilteredElementCollector(Doc)
                         .OfCategory(BuiltInCategory.OST_TitleBlocks)
                         .WhereElementIsNotElementType()
                         .Cast<FamilyInstance>())
                foreach (Parameter p in tb.Parameters)
                    names.Add(p.Definition.Name);

            // Sheet
            foreach (var vs in new FilteredElementCollector(Doc)
                         .OfClass(typeof(ViewSheet))
                         .Cast<ViewSheet>())
                foreach (Parameter p in vs.Parameters)
                    names.Add(p.Definition.Name);

            // Project bindings
            var it = Doc.ParameterBindings.ForwardIterator();
            while (it.MoveNext())
                names.Add(it.Key.Name);

            foreach (var n in names.OrderBy(n => n))
                ParamOptions.Add(n);

            // Move default lên đầu
            string[] defs = {
                "AWS Code-SP","AWS Originator-SP","AWS Zone-SP",
                "AWS Level-SP","AWS Type-SP","Sheet Number","Sheet Name"
            };
            foreach (var d in defs.Reverse())
                if (ParamOptions.Contains(d))
                    ParamOptions.Move(ParamOptions.IndexOf(d), 0);
                else
                    ParamOptions.Insert(0, d);
        }

        void InitFileNameItems()
        {
            FileNameItems.Clear();
            FileNameItems.Add(new FileNameItem("AWS Code-SP", "-"));
            FileNameItems.Add(new FileNameItem("AWS Originator-SP", "-"));
            FileNameItems.Add(new FileNameItem("AWS Zone-SP", "-"));
            FileNameItems.Add(new FileNameItem("AWS Level-SP", "-"));
            FileNameItems.Add(new FileNameItem("AWS Type-SP", "-"));
            FileNameItems.Add(new FileNameItem("Sheet Number", "_"));
            FileNameItems.Add(new FileNameItem("Sheet Name", ""));
        }

        void LoadSavedSettings()
        {
            var s = LayerExportHelper.LoadSettings();
            if (s == null) return;
            if (Directory.Exists(s.ExportPath)) ExportPath = s.ExportPath;
            if (s.Params?.Count == FileNameItems.Count &&
                s.Seps?.Count == FileNameItems.Count)
            {
                for (int i = 0; i < FileNameItems.Count; i++)
                {
                    FileNameItems[i].SelectedParam = s.Params[i];
                    FileNameItems[i].Sep = s.Seps[i];
                }
            }
        }

        void SaveCurrentSettings()
        {
            LayerExportHelper.SaveSettings(new LayerExportHelper.ExportSettings
            {
                ExportPath = ExportPath,
                Params = FileNameItems.Select(f => f.SelectedParam).ToList(),
                Seps = FileNameItems.Select(f => f.Sep).ToList()
            });
        }

        #endregion

        static void Msg(string t) => System.Windows.MessageBox.Show(t);
    }
}
