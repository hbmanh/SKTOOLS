using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SKRevitAddins.Utils;

namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
{
    public class AutoPlaceElementFrBlockCADViewModel : ViewModelBase
    {
        public Action<string> UpdateStatus { get; set; }
        public Action<RequestId> RaiseRequest { get; set; }
        public UIApplication UiApp { get; private set; }
        private UIDocument UiDoc;
        private Document ThisDoc;

        public AutoPlaceElementFrBlockCADViewModel(UIApplication uiApp, Action<RequestId> raiseRequest)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisDoc = UiDoc.ActiveView.Document;
            RaiseRequest = raiseRequest;

            var refLinkCad = UiDoc.Selection.PickObject(ObjectType.Element, new ImportInstanceSelectionFilter(), "Select Link File");
            var selectedCadLink = ThisDoc.GetElement(refLinkCad) as ImportInstance;
            if (selectedCadLink == null)
            {
                TaskDialog.Show("Error", "No valid CAD link selected.");
                return;
            }

            Level = UiDoc.ActiveView.GenLevel;
            Categories = new ObservableCollection<Category>(GetCategories(ThisDoc));
            SelectedCategory = Categories.FirstOrDefault();
            CacheFamiliesAndSymbols();

            var blocks = GeometryHelper.GetBlockNamesFromCadLink(selectedCadLink)
                .GroupBy(b => b.Block.Symbol.Name)
                .OrderBy(g => g.Key, StringComparer.CurrentCultureIgnoreCase);

            foreach (var blockGroup in blocks)
            {
                BlockMappings.Add(new BlockMapping(blockGroup, ThisDoc, Categories, Families, TypeSymbols));
            }

            // Validate lần đầu sau khi load
            ValidateAll();
        }

        public ObservableCollection<BlockMapping> BlockMappings { get; set; } = new();
        public ObservableCollection<BlockMapping> SelectedBlockMappings { get; set; } = new(); // cho batch edit

        public ObservableCollection<Category> Categories { get; set; }
        private Category _selectedCategory;
        public Category SelectedCategory
        {
            get => _selectedCategory;
            set
            {
                _selectedCategory = value;
                OnPropertyChanged(nameof(SelectedCategory));
                UpdateFamilies();
                UpdateTypeSymbols();
            }
        }

        private Level _level;
        public Level Level
        {
            get => _level;
            set
            {
                _level = value;
                OnPropertyChanged(nameof(Level));
            }
        }

        public ObservableCollection<Family> Families { get; set; } = new();
        private Family _selectedFamily;
        public Family SelectedFamily
        {
            get => _selectedFamily;
            set
            {
                _selectedFamily = value;
                OnPropertyChanged(nameof(SelectedFamily));
                UpdateTypeSymbols();
            }
        }

        public ObservableCollection<FamilySymbol> TypeSymbols { get; set; } = new();
        private FamilySymbol _selectedTypeSymbol;
        public FamilySymbol SelectedTypeSymbol
        {
            get => _selectedTypeSymbol;
            set
            {
                _selectedTypeSymbol = value;
                OnPropertyChanged(nameof(SelectedTypeSymbol));
            }
        }

        private Dictionary<ElementId, List<Family>> _categoryToFamilies = new();
        private Dictionary<ElementId, List<FamilySymbol>> _familyToSymbols = new();

        private void CacheFamiliesAndSymbols()
        {
            var families = new FilteredElementCollector(ThisDoc).OfClass(typeof(Family)).Cast<Family>();
            foreach (var fam in families)
            {
                if (fam.FamilyCategory == null) continue;
                if (!_categoryToFamilies.ContainsKey(fam.FamilyCategory.Id))
                    _categoryToFamilies[fam.FamilyCategory.Id] = new List<Family>();
                _categoryToFamilies[fam.FamilyCategory.Id].Add(fam);
            }

            var symbols = new FilteredElementCollector(ThisDoc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>();
            foreach (var sym in symbols)
            {
                if (!_familyToSymbols.ContainsKey(sym.Family.Id))
                    _familyToSymbols[sym.Family.Id] = new List<FamilySymbol>();
                _familyToSymbols[sym.Family.Id].Add(sym);
            }

            Families = new ObservableCollection<Family>(families.OrderBy(f => f.Name));
            TypeSymbols = new ObservableCollection<FamilySymbol>(symbols.OrderBy(s => s.Name));
        }

        private void UpdateFamilies()
        {
            if (SelectedCategory != null && _categoryToFamilies.TryGetValue(SelectedCategory.Id, out var fams))
            {
                Families = new ObservableCollection<Family>(fams.OrderBy(f => f.Name));
                SelectedFamily = Families.FirstOrDefault();
            }
        }

        private void UpdateTypeSymbols()
        {
            if (SelectedFamily != null && _familyToSymbols.TryGetValue(SelectedFamily.Id, out var symbols))
            {
                TypeSymbols = new ObservableCollection<FamilySymbol>(symbols.OrderBy(s => s.Name));
                SelectedTypeSymbol = TypeSymbols.FirstOrDefault();
            }
        }

        private static List<Category> GetCategories(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Select(f => f.FamilyCategory)
                .Where(c => c != null && c.CategoryType == CategoryType.Model)
                .GroupBy(c => c.Id)
                .Select(g => g.First())
                .OrderBy(c => c.Name)
                .ToList();
        }

        public void RequestPlaceElements()
        {
            RaiseRequest?.Invoke(RequestId.OK);
        }

        public void OnPropertyChanged(string propertyName) =>
            base.OnPropertyChanged(propertyName);

        public void ValidateAll()
        {
            foreach(var bm in BlockMappings)
            {
                if (!bm.IsValid)
                {
                    bm.IsEnabled = false;
                    bm.FailureNote = "Thiếu cấu hình Family/Type/Category!";
                }
                else
                {
                    bm.FailureNote = "";
                }
            }
        }

        // --- SmartSuggest và tokenizing như cũ
        public static IEnumerable<string> GetTokens(string name)
        {
            return name?.Split('_', '-', ' ', '.', '@')
                        .Select(s => s.ToLowerInvariant())
                        .Where(s => s.Length > 2)
                        ?? Enumerable.Empty<string>();
        }
        public static T SmartSuggest<T>(string blockName, IEnumerable<T> items, Func<T, string> nameSelector)
        {
            var tokens = GetTokens(blockName).ToList();
            return items
                .OrderByDescending(item => tokens.Count(t => nameSelector(item).ToLower().Contains(t)))
                .ThenBy(item => nameSelector(item))
                .FirstOrDefault();
        }

        // --- BlockMapping, gộp luôn BlockWithLink vào trong (không tạo class mới) ---
        public class BlockMapping : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;

            public Document ThisDoc { get; set; }
            public List<BlockWithLinkInfo> Blocks { get; set; }
            public int BlockCount => Blocks?.Count ?? 0;

            private bool _isEnabled = true;
            public bool IsEnabled
            {
                get => _isEnabled;
                set
                {
                    _isEnabled = value;
                    OnPropertyChanged(nameof(IsEnabled));
                    OnPropertyChanged(nameof(RowStatusColor));
                }
            }

            private int _placedCount;
            public int PlacedCount
            {
                get => _placedCount;
                set
                {
                    _placedCount = value;
                    OnPropertyChanged(nameof(PlacedCount));
                    OnPropertyChanged(nameof(RowStatusColor));
                }
            }

            private bool _hasPlacementRun;
            public bool HasPlacementRun
            {
                get => _hasPlacementRun;
                set
                {
                    _hasPlacementRun = value;
                    OnPropertyChanged(nameof(HasPlacementRun));
                    OnPropertyChanged(nameof(RowStatusColor));
                }
            }

            public string RowStatusColor
            {
                get
                {
                    if (!IsEnabled) return "LightGray";
                    if (!HasPlacementRun) return "White";
                    return PlacedCount == BlockCount ? "LightGreen" : "LightCoral";
                }
            }

            private string _blockName;
            public string BlockName
            {
                get => _blockName;
                set
                {
                    _blockName = value;
                    OnPropertyChanged(nameof(BlockName));
                    OnPropertyChanged(nameof(DisplayBlockName));
                }
            }

            public string DisplayBlockName =>
                string.IsNullOrEmpty(BlockName) ? "" :
                (BlockName.LastIndexOf('.') >= 0
                    ? BlockName.Substring(BlockName.LastIndexOf('.') + 1)
                    : BlockName);

            private string _failureNote;
            public string FailureNote
            {
                get => _failureNote;
                set
                {
                    _failureNote = value;
                    OnPropertyChanged(nameof(FailureNote));
                }
            }

            public ObservableCollection<Category> CategoriesMapping { get; set; }
            private Category _selectedCategoryMapping;
            public Category SelectedCategoryMapping
            {
                get => _selectedCategoryMapping;
                set
                {
                    if (_selectedCategoryMapping != value)
                    {
                        _selectedCategoryMapping = value;
                        OnPropertyChanged(nameof(SelectedCategoryMapping));
                        UpdateFamiliesMapping();
                        SelectedFamilyMapping = FamiliesMapping.FirstOrDefault();
                        UpdateTypeSymbolsMapping();
                        SelectedTypeSymbolMapping = TypeSymbolsMapping.FirstOrDefault();
                        // validate ngay khi thay đổi
                        OnPropertyChanged(nameof(IsValid));
                    }
                }
            }
            public ObservableCollection<Family> FamiliesMapping { get; set; }
            private Family _selectedFamilyMapping;
            public Family SelectedFamilyMapping
            {
                get => _selectedFamilyMapping;
                set
                {
                    if (_selectedFamilyMapping != value)
                    {
                        _selectedFamilyMapping = value;
                        OnPropertyChanged(nameof(SelectedFamilyMapping));
                        UpdateTypeSymbolsMapping();
                        SelectedTypeSymbolMapping = TypeSymbolsMapping.FirstOrDefault();
                        OnPropertyChanged(nameof(IsValid));
                    }
                }
            }
            public ObservableCollection<FamilySymbol> TypeSymbolsMapping { get; set; }
            private FamilySymbol _selectedTypeSymbolMapping;
            public FamilySymbol SelectedTypeSymbolMapping
            {
                get => _selectedTypeSymbolMapping;
                set
                {
                    _selectedTypeSymbolMapping = value;
                    OnPropertyChanged(nameof(SelectedTypeSymbolMapping));
                    OnPropertyChanged(nameof(IsValid));
                }
            }

            private double _offset;
            public double Offset
            {
                get => _offset;
                set
                {
                    _offset = value;
                    OnPropertyChanged(nameof(Offset));
                }
            }

            public bool IsValid =>
                SelectedTypeSymbolMapping != null
                && SelectedFamilyMapping != null
                && SelectedCategoryMapping != null;

            // Gộp BlockWithLink vào luôn (không tách class mới)
            public class BlockWithLinkInfo : INotifyPropertyChanged
            {
                public GeometryInstance Block { get; set; }
                public ImportInstance CadLink { get; set; }
                private bool _isEnabled = true;
                public bool IsEnabled
                {
                    get => _isEnabled;
                    set
                    {
                        _isEnabled = value;
                        OnPropertyChanged(nameof(IsEnabled));
                    }
                }
                public event PropertyChangedEventHandler PropertyChanged;
                public void OnPropertyChanged(string propertyName)
                {
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
                }
            }

            public BlockMapping(IGrouping<string, BlockWithLink> blocks,
                Document doc,
                ObservableCollection<Category> categories,
                ObservableCollection<Family> families,
                ObservableCollection<FamilySymbol> typeSymbols)
            {
                ThisDoc = doc;
                Blocks = blocks.Select(b => new BlockWithLinkInfo
                {
                    Block = b.Block,
                    CadLink = b.CadLink,
                    IsEnabled = true
                }).ToList();

                BlockName = blocks.Key;
                CategoriesMapping = categories;
                FamiliesMapping = families;
                TypeSymbolsMapping = typeSymbols;
                Offset = 2600;
                IsEnabled = true;

                SelectedCategoryMapping = SmartSuggest(BlockName, CategoriesMapping, c => c.Name) ?? categories.FirstOrDefault();
                UpdateFamiliesMapping();
                SelectedFamilyMapping = SmartSuggest(BlockName, FamiliesMapping, f => f.Name) ?? FamiliesMapping.FirstOrDefault();
                UpdateTypeSymbolsMapping();
                SelectedTypeSymbolMapping = SmartSuggest(BlockName, TypeSymbolsMapping, t => t.Name) ?? TypeSymbolsMapping.FirstOrDefault();
            }

            private void UpdateFamiliesMapping()
            {
                if (SelectedCategoryMapping != null && ThisDoc != null)
                {
                    var families = new FilteredElementCollector(ThisDoc)
                        .OfClass(typeof(Family))
                        .Cast<Family>()
                        .Where(f => f.FamilyCategory.Id == SelectedCategoryMapping.Id)
                        .OrderBy(f => f.Name)
                        .ToList();

                    FamiliesMapping = new ObservableCollection<Family>(families);
                    OnPropertyChanged(nameof(FamiliesMapping));
                }
            }
            private void UpdateTypeSymbolsMapping()
            {
                if (SelectedFamilyMapping != null && ThisDoc != null)
                {
                    var symbols = new FilteredElementCollector(ThisDoc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .Where(fs => fs.Family.Id == SelectedFamilyMapping.Id)
                        .OrderBy(fs => fs.Name)
                        .ToList();

                    TypeSymbolsMapping = new ObservableCollection<FamilySymbol>(symbols);
                    OnPropertyChanged(nameof(TypeSymbolsMapping));
                }
            }

            public void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
