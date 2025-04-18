using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace SKRevitAddins.Commands.DWGExport
{
    public class DWGExportViewModel : INotifyPropertyChanged
    {
        private readonly UIApplication _uiApp;
        private readonly Document _doc;
        private readonly Dictionary<string, Dictionary<string, HashSet<string>>> _catParamDic =
    new Dictionary<string, Dictionary<string, HashSet<string>>>();


        public DWGExportViewModel(UIApplication uiApp)
        {
            _uiApp = uiApp;
            _doc = uiApp.ActiveUIDocument.Document;
            Sheets = new ObservableCollection<SheetItem>();
            SelectedSheets = new ObservableCollection<SheetItem>();
            CategoryItems = new ObservableCollection<CategoryMapItem>();
            AvailableCats = new ObservableCollection<string>(_allowedCatNames);

            AddRowCmd = new RelayCommand(_ => CategoryItems.Add(new CategoryMapItem(this)));
            DeleteRowCmd = new RelayCommand(_ =>
            {
                if (SelectedMapping != null)
                    CategoryItems.Remove(SelectedMapping);
            });

            LoadSheets();
            ProgressMax = 1;
        }

        public ObservableCollection<SheetItem> Sheets { get; }
        public ObservableCollection<SheetItem> SelectedSheets { get; }
        public ObservableCollection<CategoryMapItem> CategoryItems { get; }
        public ObservableCollection<string> AvailableCats { get; }

        public ICommand AddRowCmd { get; }
        public ICommand DeleteRowCmd { get; }

        private CategoryMapItem _sel;
        public CategoryMapItem SelectedMapping
        {
            get => _sel;
            set { _sel = value; On(); }
        }

        private string _status;
        public string ExportStatusMessage
        {
            get => _status;
            set { _status = value; On(); }
        }

        private int _pVal, _pMax;
        public int ProgressValue
        {
            get => _pVal;
            set { _pVal = value; On(); }
        }
        public int ProgressMax
        {
            get => _pMax;
            set { _pMax = value; On(); }
        }

        private void LoadSheets()
        {
            var sheets = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .OrderBy(s => s.SheetNumber);
            foreach (var sh in sheets)
                Sheets.Add(new SheetItem { Sheet = sh });
        }

        public void RefreshCategoriesFromSelectedSheets()
        {
            var view = System.Windows.Data.CollectionViewSource.GetDefaultView(CategoryItems);
            using (view.DeferRefresh())
            {
                CategoryItems.Clear();
                _catParamDic.Clear();
                if (!SelectedSheets.Any()) return;

                var rows = new Dictionary<(string, string, string), CategoryMapItem>();
                var firstRowOfCat = new Dictionary<string, CategoryMapItem>();

                foreach (var si in SelectedSheets)
                    CollectFromSheet(si.Sheet, rows, firstRowOfCat);

                foreach (var row in rows.Values
                                         .OrderBy(r => r.CategoryName)
                                         .ThenBy(r => r.ParameterName)
                                         .ThenBy(r => r.ParamValue))
                {
                    CategoryItems.Add(row);
                }
            }
        }

        private void CollectFromSheet(ViewSheet sheet,
                                      IDictionary<(string, string, string), CategoryMapItem> rows,
                                      IDictionary<string, CategoryMapItem> firstRowOfCat)
        {
            Gather(sheet, BuiltInCategory.OST_TitleBlocks, rows, firstRowOfCat);

            var vps = new FilteredElementCollector(_doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .Where(vp => vp.SheetId == sheet.Id);

            foreach (var vp in vps)
            {
                var v = _doc.GetElement(vp.ViewId) as View;
                if (v != null)
                    Gather(v, null, rows, firstRowOfCat);
            }
        }

        private ElementFilter _catFilter;
        private void Gather(Element viewOrSheet, BuiltInCategory? bic,
                            IDictionary<(string, string, string), CategoryMapItem> rows,
                            IDictionary<string, CategoryMapItem> firstRowOfCat)
        {
            if (_catFilter == null)
            {
                var flt = _allowedCats
                    .Select(c => new ElementCategoryFilter(c))
                    .Cast<ElementFilter>()
                    .ToList();
                _catFilter = new LogicalOrFilter(flt);
            }

            var col = new FilteredElementCollector(_doc, viewOrSheet.Id)
                          .WherePasses(_catFilter)
                          .WhereElementIsNotElementType();

            foreach (Element el in col)
            {
                if (el is ImportInstance || el.Category == null) continue;

                var bicId = (BuiltInCategory)el.Category.Id.IntegerValue;
                if (!_allowedCats.Contains(bicId)) continue;

                string cat = el.Category.Name;

                foreach (Parameter p in el.Parameters)
                {
                    if (p.StorageType != StorageType.String || !p.HasValue) continue;

                    string paramName = p.Definition.Name;
                    string paramVal = p.AsValueString();
                    if (string.IsNullOrEmpty(paramVal))
                        paramVal = p.AsString() ?? "";
                    if (string.IsNullOrEmpty(paramVal)) continue;

                    if (!_catParamDic.TryGetValue(cat, out var pMap))
                        pMap = _catParamDic[cat] = new Dictionary<string, HashSet<string>>();

                    if (!pMap.TryGetValue(paramName, out var set))
                        set = pMap[paramName] = new HashSet<string>();

                    if (set.Add(paramVal) && firstRowOfCat.TryGetValue(cat, out var first))
                    {
                        if (!first.ParameterOptions.Contains(paramName))
                            first.ParameterOptions.Add(paramName);
                        if (first.ParameterName == paramName &&
                            !first.ParamOptions.Contains(paramVal))
                            first.ParamOptions.Add(paramVal);
                    }

                    var key = (cat, paramName, paramVal);
                    if (rows.ContainsKey(key)) continue;

                    var row = new CategoryMapItem(this)
                    {
                        CategoryName = cat,
                        ParameterName = paramName,
                        ParamValue = paramVal,
                        LayerName = cat.Replace(' ', '_'),
                        ColorIndex = 7
                    };
                    row.ParameterOptions.Add(paramName);
                    row.ParamOptions.Add(paramVal);

                    rows[key] = row;
                    if (!firstRowOfCat.ContainsKey(cat))
                        firstRowOfCat[cat] = row;
                }
            }
        }

        internal void UpdateParameterList(CategoryMapItem row)
        {
            row.ParameterOptions.Clear();
            row.ParamOptions.Clear();

            if (row.CategoryName != null &&
                _catParamDic.TryGetValue(row.CategoryName, out var pMap))
            {
                foreach (var p in pMap.Keys) row.ParameterOptions.Add(p);
                row.ParameterName = row.ParameterOptions.FirstOrDefault();
            }
            else
            {
                row.ParameterName = "";
                row.ParamValue = "";
            }

            row.On(nameof(row.ParameterOptions));
            row.On(nameof(row.ParameterName));
            UpdateParamValues(row);
        }

        internal void UpdateParamValues(CategoryMapItem row)
        {
            row.ParamOptions.Clear();

            if (row.CategoryName != null &&
                _catParamDic.TryGetValue(row.CategoryName, out var pMap) &&
                row.ParameterName != null &&
                pMap.TryGetValue(row.ParameterName, out var set))
            {
                foreach (var v in set) row.ParamOptions.Add(v);
                if (!set.Contains(row.ParamValue))
                    row.ParamValue = row.ParamOptions.FirstOrDefault() ?? "";
            }
            else
            {
                row.ParamValue = "";
            }

            row.On(nameof(row.ParamOptions));
            row.On(nameof(row.ParamValue));
        }

        private static readonly HashSet<BuiltInCategory> _allowedCats = new()
        {
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_PlaceHolderDucts,
            BuiltInCategory.OST_PlaceHolderPipes,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_TitleBlocks
        };

        private static readonly List<string> _allowedCatNames = _allowedCats
            .Select(LabelUtils.GetLabelFor)
            .Distinct()
            .ToList();

        public event PropertyChangedEventHandler PropertyChanged;
        private void On([CallerMemberName] string prop = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }

    public class SheetItem
    {
        public ViewSheet Sheet { get; set; }
        public string SheetDisplay => $"{Sheet.SheetNumber}-{Sheet.Name}";
    }

    public class CategoryMapItem : INotifyPropertyChanged
    {
        private readonly DWGExportViewModel _owner;
        public CategoryMapItem(DWGExportViewModel owner) { _owner = owner; }

        private string _cat, _paramName, _paramValue, _layer;
        private int _color = 7;

        public string CategoryName
        {
            get => _cat;
            set
            {
                if (_cat == value) return;
                _cat = value; On();
                _owner.UpdateParameterList(this);
            }
        }

        public string ParameterName
        {
            get => _paramName;
            set
            {
                if (_paramName == value) return;
                _paramName = value; On();
                _owner.UpdateParamValues(this);
            }
        }

        public string ParamValue
        {
            get => _paramValue;
            set { _paramValue = value; On(); }
        }

        public string LayerName
        {
            get => _layer;
            set { _layer = value; On(); }
        }

        public int ColorIndex
        {
            get => _color;
            set { _color = value; On(); }
        }

        public ObservableCollection<string> ParameterOptions { get; } = new();
        public ObservableCollection<string> ParamOptions { get; } = new();

        public event PropertyChangedEventHandler PropertyChanged;
        public void On([CallerMemberName] string p = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
    }
}