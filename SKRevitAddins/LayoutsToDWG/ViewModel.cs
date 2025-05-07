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
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

namespace SKRevitAddins.LayoutsToDWG.ViewModel
{
    public class InvertBoolConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object v, Type t, object p, System.Globalization.CultureInfo c) => !(bool)v;
        public object ConvertBack(object v, Type t, object p, System.Globalization.CultureInfo c) => !(bool)v;
    }

    public class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    public class SheetItem : ViewModelBase
    {
        public SheetItem(ViewSheet s) => Sheet = s;
        public ViewSheet Sheet { get; }
        public string SheetNumber => Sheet.SheetNumber;
        public string SheetName => Sheet.Name;
        bool _sel = true;
        public bool IsSelected { get => _sel; set { _sel = value; OnPropertyChanged(); } }
    }

    public class FileNameItem : ViewModelBase
    {
        public FileNameItem(string param, string sep)
        { SelectedParam = param; Sep = sep; }
        public string SelectedParam { get => _param; set { _param = value; OnPropertyChanged(); } }
        public string Sep { get => _sep; set { _sep = value; OnPropertyChanged(); } }
        string _param, _sep;
    }

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

    public class LayoutsToDWGViewModel : ViewModelBase
    {
        readonly UIDocument _uiDoc;
        readonly ExportSheetsHandler _handler;
        readonly ExternalEvent _evt;
        Document Doc => _uiDoc.Document;

        public LayoutsToDWGViewModel(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;
            _handler = new ExportSheetsHandler();
            _evt = ExternalEvent.Create(_handler);

            LoadExportSetups();
            LoadSheetSets();
            LoadParamOptions();
            InitFileNameItems();
            LoadSavedSettings();

            BrowseFolderCmd = new RelayCommand(_ => BrowseFolder(), _ => !IsBusy);
            ExportLayerCmd = new RelayCommand(_ => ExportLayerMapping(), _ => !IsBusy);
            StartExportSheetsCmd = new RelayCommand(_ => StartExport(), _ => !IsBusy);
            CancelExportCmd = new RelayCommand(_ => CancelExport(), _ => IsBusy);
            CancelCmd = new RelayCommand(_ => RequestClose?.Invoke());
        }

        public ObservableCollection<string> ExportSetups { get; } = new();
        public ObservableCollection<string> SheetSets { get; } = new();
        public ObservableCollection<SheetItem> SheetItems { get; } = new();
        public ObservableCollection<FileNameItem> FileNameItems { get; } = new();
        public ObservableCollection<string> ParamOptions { get; } = new();

        public string SelectedExportSetup { get; set; }

        string _exportPath;
        public string ExportPath { get => _exportPath; set { _exportPath = value; OnPropertyChanged(); } }

        string _selSet;
        public string SelectedSheetSet
        {
            get => _selSet;
            set { _selSet = value; OnPropertyChanged(); LoadSheetsFromSet(); }
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

        public ICommand BrowseFolderCmd { get; }
        public ICommand ExportLayerCmd { get; }
        public ICommand StartExportSheetsCmd { get; }
        public ICommand CancelCmd { get; }
        public ICommand CancelExportCmd { get; }

        public Action RequestClose { get; set; }

        void CancelExport()
        {
            _handler.IsCancelled = true;
            Msg("Cancelling export...");
        }

        void LoadParamOptions()
        {
            ParamOptions.Clear();
            var paramNames = new HashSet<string>();

            var titleBlocks = new FilteredElementCollector(Doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>();

            foreach (var tb in titleBlocks)
                foreach (Parameter p in tb.Parameters)
                    paramNames.Add(p.Definition.Name);

            var sheets = new FilteredElementCollector(Doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>();

            foreach (var sheet in sheets)
                foreach (Parameter p in sheet.Parameters)
                    paramNames.Add(p.Definition.Name);

            var bindingMap = Doc.ParameterBindings;
            var iterator = bindingMap.ForwardIterator();
            while (iterator.MoveNext())
            {
                var definition = iterator.Key;
                paramNames.Add(definition.Name);
            }

            foreach (var name in paramNames.OrderBy(n => n))
                ParamOptions.Add(name);

            string[] def = {
                "AWS Code-SP", "AWS Originator-SP", "AWS Zone-SP",
                "AWS Level-SP", "AWS Type-SP", "Sheet Number", "Sheet Name" };

            foreach (string d in def.Reverse())
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

        string GetParamValue(ViewSheet sheet, string paramName)
        {
            var param = sheet.LookupParameter(paramName);
            if (param != null && param.HasValue)
                return param.AsString();

            var tb = new FilteredElementCollector(Doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .FirstOrDefault(f => f.OwnerViewId == sheet.Id);

            if (tb?.LookupParameter(paramName) is Parameter tbParam && tbParam.HasValue)
                return tbParam.AsString();

            var projParam = new FilteredElementCollector(Doc)
                .WhereElementIsNotElementType()
                .Where(e => e.GetTypeId() != ElementId.InvalidElementId)
                .Select(e => e.LookupParameter(paramName))
                .FirstOrDefault(p => p != null && p.HasValue);

            if (projParam != null)
                return projParam.AsString();

            var sharedParam = new FilteredElementCollector(Doc)
                .WhereElementIsNotElementType()
                .SelectMany(e => e.Parameters.Cast<Parameter>())
                .FirstOrDefault(p => p.Definition.Name == paramName && p.IsShared && p.HasValue);

            return sharedParam?.AsString() ?? "";
        }

        string BuildFileName(ViewSheet sheet)
        {
            var fileNameParts = new List<string>();
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
                    // Ưu tiên lấy từ sheet trước, sau đó mới đến title block
                    var pSheet = sheet.LookupParameter(item.SelectedParam);
                    if (pSheet != null && pSheet.HasValue)
                        value = pSheet.AsString();
                    else
                        value = GetParamValue(sheet, item.SelectedParam); // GetParamValue đã xử lý title block
                }

                // Sanitize giá trị
                value = LayerExportHelper.Sanitize(value?.Trim() ?? "");

                // Chỉ thêm phần này vào tên file nếu nó có giá trị
                if (!string.IsNullOrWhiteSpace(value))
                {
                    fileNameParts.Add(value);
                }
                // Nếu phần hiện tại có giá trị VÀ không phải là phần tử cuối cùng có giá trị
                // VÀ separator cũng có giá trị, thì thêm separator
                // Điều này tránh thêm separator thừa ở cuối hoặc giữa các phần rỗng
                if (!string.IsNullOrWhiteSpace(value) && 
                    !string.IsNullOrWhiteSpace(item.Sep) && 
                    i < FileNameItems.Count - 1 && // Không phải là item cuối cùng trong cấu hình
                    FileNameItems.Skip(i + 1).Any(nextItem => !string.IsNullOrWhiteSpace(GetParamValueOrSheetProperty(sheet, nextItem.SelectedParam)))) // Kiểm tra xem có item nào sau nó có giá trị không
                {
                    fileNameParts.Add(item.Sep);
                }
            }
            
            // Nối các phần lại và trim các ký tự phân tách thừa ở cuối nếu có
            string finalName = string.Join("", fileNameParts);
            
            // Loại bỏ các ký tự phân tách có thể bị thừa ở cuối cùng
            // sau khi đã ghép nối tất cả các phần có giá trị.
            if (FileNameItems.Any())
            {
                var lastConfiguredSeparator = FileNameItems.LastOrDefault(fni => !string.IsNullOrWhiteSpace(fni.Sep))?.Sep;
                if (!string.IsNullOrWhiteSpace(lastConfiguredSeparator) && finalName.EndsWith(lastConfiguredSeparator))
                {
                    finalName = finalName.Substring(0, finalName.Length - lastConfiguredSeparator.Length);
                }
            }
            // Hoặc một cách đơn giản hơn để trim nhiều loại separator phổ biến:
            // return finalName.TrimEnd('-', '_', ' '); // Điều chỉnh các ký tự cần trim

            return finalName;
        }

        // Hàm phụ trợ để lấy giá trị cho việc kiểm tra lookahead
        string GetParamValueOrSheetProperty(ViewSheet sheet, string paramName)
        {
            if (paramName == "Sheet Number") return sheet.SheetNumber;
            if (paramName == "Sheet Name") return sheet.LookupParameter("Sheet Name")?.AsString() ?? sheet.Name;
            
            var pSheet = sheet.LookupParameter(paramName);
            if (pSheet != null && pSheet.HasValue) return pSheet.AsString();
            
            return GetParamValue(sheet, paramName); // Hàm GetParamValue của bạn
        }

        void StartExport()
        {
            if (IsBusy) return;
            var sel = SheetItems.Where(s => s.IsSelected).ToList();
            if (!sel.Any()) { Msg("No sheet selected."); return; }
            if (!Directory.Exists(ExportPath)) { Msg("Path invalid."); return; }

            var fileNames = sel.ToDictionary(s => s.Sheet.Id, s => BuildFileName(s.Sheet));
            if (!CheckDuplicates(fileNames, ExportPath, out string warn))
            { Msg(warn); return; }

            var setup = new FilteredElementCollector(Doc)
                          .OfClass(typeof(ExportDWGSettings))
                          .Cast<ExportDWGSettings>()
                          .FirstOrDefault(x => x.Name == SelectedExportSetup);
            if (setup == null) { Msg("Export setup not found."); return; }

            var opt = setup.GetDWGExportOptions(); opt.MergedViews = true;

            _handler.ViewIds = sel.Select(s => s.Sheet.Id).ToList();
            _handler.TargetPath = ExportPath;
            _handler.Options = opt;
            _handler.FileNames = fileNames;
            _handler.OpenFolder = OpenFolderAfterExport;
            _handler.BusySetter = v => IsBusy = v;
            _handler.ProgressReporter = (cur, total) =>
            {
                Progress = total == 0 ? 0 : (double)cur / total;
            };

            IsBusy = true;
            _evt.Raise();
            SaveCurrentSettings();
        }

        static bool CheckDuplicates(Dictionary<ElementId, string> fileNames, string dir, out string msg)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in fileNames.Values)
            {
                string file = name + ".dwg";
                if (!names.Add(file)) { msg = $"Duplicate: {file}"; return false; }
                if (File.Exists(Path.Combine(dir, file))) { msg = $"Exists: {file}"; return false; }
            }
            msg = null; return true;
        }

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
                               .Select(s => s.Name)
                               .OrderBy(n => n))
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
            var comparer = new NaturalSheetComparer();

            foreach (var vs in set.Views.OfType<ViewSheet>()
                        .OrderBy(v => v.SheetNumber, comparer))
                SheetItems.Add(new SheetItem(vs));
        }

        void BrowseFolder()
        {
            using var dlg = new System.Windows.Forms.FolderBrowserDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                ExportPath = dlg.SelectedPath;
        }

        void ExportLayerMapping()
        {
            if (string.IsNullOrWhiteSpace(SelectedExportSetup))
            {
                TaskDialog.Show("Layer", "Choose export setup.");
                return;
            }

            var dlg = new SaveFileDialog { Filter = "TXT|*.txt", FileName = $"{SelectedExportSetup}.txt" };
            if (dlg.ShowDialog() != true) return;

            try
            {
                LayerExportHelper.WriteLayerMapping(Doc, SelectedExportSetup, dlg.FileName);
                TaskDialog.Show("Completed", dlg.FileName);
            }
            catch (Exception ex) { TaskDialog.Show("Error", ex.Message); }
        }

        void LoadSavedSettings()
        {
            var s = LayerExportHelper.LoadSettings();
            if (s == null) return;
            if (Directory.Exists(s.ExportPath)) ExportPath = s.ExportPath;

            if (s.Params != null && s.Seps != null &&
                s.Params.Count == FileNameItems.Count &&
                s.Seps.Count == FileNameItems.Count)
            {
                for (int i = 0; i < FileNameItems.Count; i++)
                {
                    FileNameItems[i].SelectedParam = s.Params[i];
                    FileNameItems[i].Sep = s.Seps[i];
                }
            }
        }

        public void SaveCurrentSettings()
            => LayerExportHelper.SaveSettings(new LayerExportHelper.ExportSettings
            {
                ExportPath = ExportPath,
                Params = FileNameItems.Select(r => r.SelectedParam).ToList(),
                Seps = FileNameItems.Select(r => r.Sep).ToList()
            });

        static void Msg(string t) => MessageBox.Show(t);
    }
}
