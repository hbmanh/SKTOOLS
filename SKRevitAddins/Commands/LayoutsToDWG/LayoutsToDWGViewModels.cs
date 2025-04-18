using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace LayoutsToDWG.Commands
{
    public class LayoutsToDWGViewModel : INotifyPropertyChanged
    {
        readonly UIApplication _app;
        readonly Document _doc;
        public LayoutsToDWGViewModel(UIApplication app)
        {
            _app = app;
            _doc = app.ActiveUIDocument.Document;

            ExportSetups = new ObservableCollection<string>();
            TitleblockParams = new ObservableCollection<string>();
            AllSheets = new ObservableCollection<SheetSelectionItem>();
            SelectedSheets = new List<ViewSheet>();

            LoadExportSetups();
            LoadTitleblockParams();
            LoadSheetSets();
            LoadSheets();
        }

        // ==== Settings ====
        public LayoutsToDWGRequest Request { get; } = new LayoutsToDWGRequest();

        private string _exportFolder;
        public string ExportFolder
        {
            get => _exportFolder;
            set { _exportFolder = value; On(); On(nameof(IsExportFolderValid)); }
        }
        public bool IsExportFolderValid =>
            !string.IsNullOrWhiteSpace(ExportFolder) && Directory.Exists(ExportFolder);

        private bool _openAfter = false;
        public bool OpenAfterExport
        {
            get => _openAfter;
            set { _openAfter = value; On(); }
        }

        private string _statusMessage;
        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; On(); }
        }

        private DWGExportMode _mode = DWGExportMode.SingleFile;
        public DWGExportMode ExportMode
        {
            get => _mode;
            set { _mode = value; On(); }
        }

        // ==== ExportSetups & TitleblockParams ====
        public ObservableCollection<string> ExportSetups { get; }
        private string _selectedExportSetup;
        public string SelectedExportSetup
        {
            get => _selectedExportSetup;
            set { _selectedExportSetup = value; On(); }
        }

        public ObservableCollection<string> TitleblockParams { get; }
        private string _selectedNumberParam;
        public string SelectedNumberParam
        {
            get => _selectedNumberParam;
            set { _selectedNumberParam = value; On(); }
        }
        private string _selectedNameParam;
        public string SelectedNameParam
        {
            get => _selectedNameParam;
            set { _selectedNameParam = value; On(); }
        }
        private string _selectedScaleParam;
        public string SelectedScaleParam
        {
            get => _selectedScaleParam;
            set { _selectedScaleParam = value; On(); }
        }

        // ==== Sheet selection ====
        public ObservableCollection<string> SheetSets { get; } = new();
        private string _selSheetSet;
        public string SelectedSheetSet
        {
            get => _selSheetSet;
            set { _selSheetSet = value; On(); LoadSheetSet(value); }
        }

        public ObservableCollection<SheetSelectionItem> AllSheets { get; }
        public List<ViewSheet> SelectedSheets { get; }

        public ObservableCollection<string> SortParameters { get; }
            = new ObservableCollection<string> { "Sheet Name", "Sheet Number", "Level" };
        private string _selSortParam;
        public string SelectedSortParameter
        {
            get => _selSortParam;
            set { _selSortParam = value; On(); SortSheetList(); }
        }

        // ==== Commands ====
        public System.Windows.Input.ICommand BrowseFolderCmd =>
            new RelayCommand(_ =>
            {
                using var dlg = new FolderBrowserDialog { SelectedPath = ExportFolder };
                if (dlg.ShowDialog() == DialogResult.OK)
                    ExportFolder = dlg.SelectedPath;
            });

        // ==== Load methods ====
        public void LoadExportSetups()
        {
            foreach (var name in _app.Application.ExportOptions.GetDWGExportSetupNames())
                ExportSetups.Add(name);
            SelectedExportSetup = ExportSetups.FirstOrDefault();
        }
        public void LoadTitleblockParams()
        {
            var titles = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.FamilyCategory.Id.IntegerValue
                             == (int)BuiltInCategory.OST_TitleBlocks)
                .SelectMany(fi => fi.Symbol.Parameters.Cast<Parameter>())
                .Select(p => p.Definition.Name)
                .Distinct();
            foreach (var n in titles) TitleblockParams.Add(n);
            SelectedNumberParam = TitleblockParams.ElementAtOrDefault(0);
            SelectedNameParam = TitleblockParams.ElementAtOrDefault(1);
            SelectedScaleParam = TitleblockParams.ElementAtOrDefault(2);
        }

        public void LoadSheetSets()
        {
            SheetSets.Clear();
            SheetSets.Add("-By Default-");
            foreach (ViewSheetSet s in new FilteredElementCollector(_doc)
                     .OfClass(typeof(ViewSheetSet)))
                SheetSets.Add(s.Name);
            SelectedSheetSet = SheetSets[0];
        }
        public void LoadSheets() => LoadSheetSet(SelectedSheetSet);
        public void LoadSheetSet(string name)
        {
            AllSheets.Clear();
            IEnumerable<ViewSheet> sheets;
            if (name == "-By Default-")
                sheets = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>();
            else
            {
                var set = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ViewSheetSet))
                    .Cast<ViewSheetSet>()
                    .FirstOrDefault(s => s.Name == name);
                sheets = set == null ? Enumerable.Empty<ViewSheet>()
                       : set.Views.Select(id => _doc.GetElement(id) as ViewSheet);
            }

            foreach (var sheet in sheets)
                AllSheets.Add(new SheetSelectionItem(sheet));
        }

        private void SortSheetList()
        {
            if (string.IsNullOrEmpty(SelectedSortParameter)) return;
            var sorted = SelectedSortParameter switch
            {
                "Sheet Name" => AllSheets.OrderBy(s => s.SheetName),
                "Sheet Number" => AllSheets.OrderBy(s => s.Sheet.SheetNumber),
                "Level" => AllSheets.OrderBy(s => s.Level),
                _ => AllSheets
            };
            AllSheets.Clear();
            foreach (var s in sorted) AllSheets.Add(s);
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void On([CallerMemberName] string p = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
    }

    public class SheetSelectionItem : INotifyPropertyChanged
    {
        public ViewSheet Sheet { get; }
        public SheetSelectionItem(ViewSheet sheet)
        {
            Sheet = sheet;
            SheetName = $"{sheet.SheetNumber} - {sheet.Name}";
            Level = sheet.LookupParameter("Level")?.AsString() ?? "";
            try
            {
                var box = sheet.Outline;
                var w = UnitUtils.ConvertFromInternalUnits(box.Max.U - box.Min.U, UnitTypeId.Millimeters);
                var h = UnitUtils.ConvertFromInternalUnits(box.Max.V - box.Min.V, UnitTypeId.Millimeters);
                PaperSize = $"{(int)w}x{(int)h}";
            }
            catch { PaperSize = ""; }
        }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; On(); }
        }

        public string SheetName { get; }
        public string Level { get; }
        public string PaperSize { get; }

        public event PropertyChangedEventHandler PropertyChanged;
        private void On([CallerMemberName] string p = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
    }
}
