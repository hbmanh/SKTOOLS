using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using SKRevitAddins.Utils;
using Microsoft.VisualBasic;


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
            ImportLayerSettingsCmd = new RelayCommand(_ => ImportLayerSettings());

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
        public ICommand ImportLayerSettingsCmd { get; }

        void ExportLayerSettings()
        {
            if (string.IsNullOrWhiteSpace(SelectedExportSetup))
            {
                TaskDialog.Show("Export Layer Settings", "Vui lòng chọn một Export Setup.");
                return;
            }

            // Chọn nơi lưu file
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save Layer Mapping",
                Filter = "Text File (*.txt)|*.txt",
                FileName = $"LayerMapping_{SelectedExportSetup}.txt"
            };

            if (dlg.ShowDialog() != true)
                return;

            try
            {
                LayerExportHelper.ExportLayerMappingToTxt(Doc, SelectedExportSetup, dlg.FileName);
                TaskDialog.Show("Export Completed", $"Đã xuất bảng layer thành công:\n{dlg.FileName}");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        void ImportLayerSettings()
        {
            string setupName = SelectedExportSetup;

            // Nếu chưa chọn, hỏi người dùng đặt tên
            if (string.IsNullOrWhiteSpace(setupName))
            {
                var input = Microsoft.VisualBasic.Interaction.InputBox(
                    "Bạn chưa chọn Export Setup. Vui lòng nhập tên để tạo mới:",
                    "Tạo Export Setup mới",
                    "MyExportSetup");

                if (string.IsNullOrWhiteSpace(input))
                    return;

                setupName = input;

                using (Transaction tx = new Transaction(Doc, "Create DWG Export Setup"))
                {
                    tx.Start();

                    var existing = new FilteredElementCollector(Doc)
                        .OfClass(typeof(ExportDWGSettings))
                        .Cast<ExportDWGSettings>()
                        .FirstOrDefault(x => x.Name == setupName);

                    if (existing == null)
                        ExportDWGSettings.Create(Doc, setupName);

                    tx.Commit();
                }

                // 👉 REFRESH danh sách ExportSetups sau khi tạo mới
                ExportSetups.Clear();
                foreach (var s in new FilteredElementCollector(Doc)
                                    .OfClass(typeof(ExportDWGSettings))
                                    .Cast<ExportDWGSettings>()
                                    .Select(x => x.Name)
                                    .OrderBy(n => n))
                    ExportSetups.Add(s);

                SelectedExportSetup = setupName;
            }

            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Chọn file .txt để import layer mapping",
                Filter = "Layer Mapping (*.txt)|*.txt",
                DefaultExt = ".txt"
            };

            if (dlg.ShowDialog() != true)
                return;

            string filePath = dlg.FileName;

            // 👉 VALIDATE định dạng file
            if (!LayerExportHelper.IsValidLayerMappingFile(filePath))
            {
                TaskDialog.Show("Lỗi", "File layer mapping không hợp lệ. Vui lòng kiểm tra định dạng.");
                return;
            }

            try
            {
                var dwgSettings = new FilteredElementCollector(Doc)
                    .OfClass(typeof(ExportDWGSettings))
                    .Cast<ExportDWGSettings>()
                    .FirstOrDefault(x => x.Name == SelectedExportSetup);

                if (dwgSettings == null)
                {
                    TaskDialog.Show("Error", $"Không tìm thấy export setup: {SelectedExportSetup}");
                    return;
                }

                using (Transaction tx = new Transaction(Doc, "Import Layer Mapping"))
                {
                    tx.Start();

                    var dwgOptions = dwgSettings.GetDWGExportOptions();

                    ((BaseExportOptions)dwgOptions).LayerMapping = filePath;

                    var layerTable = dwgOptions.GetExportLayerTable();
                    ((BaseExportOptions)dwgOptions).SetExportLayerTable(layerTable);

                    dwgSettings.SetDWGExportOptions(dwgOptions);

                    tx.Commit();
                }

                TaskDialog.Show("Import Layer Settings", "Đã import layer mapping thành công.");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Lỗi", ex.Message);
            }
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

            foreach (var view in set.Views)
            {
                if (view is ViewSheet vs)
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
