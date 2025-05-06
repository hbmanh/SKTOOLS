// LayoutsToDWGViewModel.cs
//----------------------------------------------------------------
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
        public bool IsSelected
        {
            get => _isSel;
            set { _isSel = value; OnPropertyChanged(); }
        }

        int _order;
        public int Order
        {
            get => _order;
            set { _order = value; OnPropertyChanged(); }
        }
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

            // DWG Export setups hiện có
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

            //── Mặc định Middle/Ending ──────────────────────────
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

        //──────── ICommand: Export / Import Layer Settings ──────
        public ICommand ExportLayerSettingsCmd { get; }
        public ICommand ImportLayerSettingsCmd { get; }

        //=================================================================
        //  A. EXPORT layer‑mapping  (file = <Tên Setup>.txt)
        //=================================================================
        void ExportLayerSettings()
        {
            if (string.IsNullOrWhiteSpace(SelectedExportSetup))
            {
                TaskDialog.Show("Export Layer Settings",
                                "Vui lòng chọn một Export Setup.");
                return;
            }

            var dlg = new SaveFileDialog
            {
                Title = "Save Layer Mapping",
                Filter = "Text File (*.txt)|*.txt",
                FileName = $"{SelectedExportSetup}.txt"          // CHANGED
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                LayerExportHelper.ExportLayerMappingToTxt(
                    Doc, SelectedExportSetup, dlg.FileName);

                TaskDialog.Show("Export Completed",
                                $"Đã xuất bảng layer thành công:\n{dlg.FileName}");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        //=================================================================
        //  B. IMPORT layer‑mapping
        //     - setup tồn tại ➜ gán file .txt
        //     - setup chưa có ➜ tạo mới tên = tên file rồi gán
        //=================================================================
        void ImportLayerSettings()
        {
            // ➊ Chọn file trước
            var dlg = new OpenFileDialog
            {
                Title = "Chọn file .txt để import layer mapping",
                Filter = "Layer Mapping (*.txt)|*.txt",
                DefaultExt = ".txt"
            };
            if (dlg.ShowDialog() != true) return;

            string filePath = dlg.FileName;
            string setupName = Path.GetFileNameWithoutExtension(filePath);

            // ➋ Validate file
            if (!LayerExportHelper.IsValidLayerMappingFile(filePath))
            {
                TaskDialog.Show("Lỗi", "File layer‑mapping không hợp lệ.");
                return;
            }

            // ➌ Tìm hoặc tạo ExportDWGSettings
            ExportDWGSettings dwgSettings = new FilteredElementCollector(Doc)
                .OfClass(typeof(ExportDWGSettings))
                .Cast<ExportDWGSettings>()
                .FirstOrDefault(x => x.Name == setupName);

            using (var tx = new Transaction(Doc, "Import Layer Mapping"))
            {
                tx.Start();

                if (dwgSettings == null)
                    dwgSettings = ExportDWGSettings.Create(Doc, setupName);

                var opt = dwgSettings.GetDWGExportOptions();
                ((BaseExportOptions)opt).LayerMapping = filePath;   // gán file
                dwgSettings.SetDWGExportOptions(opt);

                tx.Commit();
            }

            // ➍ Cập nhật UI
            if (!ExportSetups.Contains(setupName))
                ExportSetups.Add(setupName);
            SelectedExportSetup = setupName;

            TaskDialog.Show("Import Layer Settings",
                            $"Đã import layer mapping cho setup “{setupName}” thành công.");
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
