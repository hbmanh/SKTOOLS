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
    public class BlockWithLink
    {
        public GeometryInstance Block { get; set; }
        public ImportInstance CadLink { get; set; }
    }

    public class AutoPlaceElementFrBlockCADViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Document ThisDoc;

        public AutoPlaceElementFrBlockCADViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisDoc = UiDoc.ActiveView.Document;

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

            UpdateFamilies();
            UpdateTypeSymbols();

            var blocks = GetBlockNamesFromCadLink(selectedCadLink)
                .GroupBy(b => b.Block.Symbol.Name)
                .OrderBy(g => g.Key, System.StringComparer.CurrentCultureIgnoreCase);

            foreach (var blockGroup in blocks)
                BlockMappings.Add(new BlockMapping(blockGroup, ThisDoc, Categories, Families, TypeSymbols));
        }

        public ObservableCollection<BlockMapping> BlockMappings { get; set; } = new ObservableCollection<BlockMapping>();

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

        public ObservableCollection<Family> Families { get; set; } = new ObservableCollection<Family>();

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

        public ObservableCollection<FamilySymbol> TypeSymbols { get; set; } = new ObservableCollection<FamilySymbol>();

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

        private void UpdateFamilies()
        {
            if (SelectedCategory != null)
            {
                Families = new ObservableCollection<Family>(
                    new FilteredElementCollector(ThisDoc)
                        .OfClass(typeof(Family))
                        .Cast<Family>()
                        .Where(family => family.FamilyCategory.Id == SelectedCategory.Id)
                        .OrderBy(family => family.Name));
                SelectedFamily = Families.FirstOrDefault();
            }
        }

        private void UpdateTypeSymbols()
        {
            if (SelectedFamily != null)
            {
                TypeSymbols = new ObservableCollection<FamilySymbol>(
                    new FilteredElementCollector(ThisDoc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .Where(fs => fs.Family.Id == SelectedFamily.Id)
                        .OrderBy(fs => fs.Name));
                SelectedTypeSymbol = TypeSymbols.FirstOrDefault();
            }
        }

        private static List<Category> GetCategories(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Select(family => family.FamilyCategory)
                .Where(c => c != null && c.CategoryType == CategoryType.Model)
                .GroupBy(c => c.Id)
                .Select(g => g.First())
                .OrderBy(c => c.Name)
                .ToList();
        }

        private static List<BlockWithLink> GetBlockNamesFromCadLink(ImportInstance cadLink)
        {
            var blocks = new List<BlockWithLink>();
            GeometryElement geoElement = cadLink.get_Geometry(new Options());

            foreach (GeometryObject geoObject in geoElement)
            {
                if (geoObject is GeometryInstance instance)
                {
                    foreach (GeometryObject instObj in instance.SymbolGeometry)
                    {
                        if (instObj is GeometryInstance blockInstance)
                            blocks.Add(new BlockWithLink { Block = blockInstance, CadLink = cadLink });
                    }
                }
            }
            return blocks;
        }
    }

    public class BlockMapping : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public Document ThisDoc { get; set; }
        public List<BlockWithLink> Blocks { get; set; }
        public int BlockCount => Blocks?.Count ?? 0;

        private bool _isEnabled = true;
        public bool IsEnabled
        {
            get => _isEnabled;
            set { _isEnabled = value; OnPropertyChanged(nameof(IsEnabled)); }
        }

        public BlockMapping(IGrouping<string, BlockWithLink> blocks, Document doc, ObservableCollection<Category> categories, ObservableCollection<Family> families, ObservableCollection<FamilySymbol> typeSymbols)
        {
            ThisDoc = doc;
            Blocks = blocks.ToList();
            BlockName = blocks.First().Block.Symbol.Name;
            CategoriesMapping = categories;
            FamiliesMapping = families;
            TypeSymbolsMapping = typeSymbols;
            Offset = 2600;
            IsEnabled = true;

            SelectedCategoryMapping = SmartSuggestCategory(DisplayBlockName, CategoriesMapping) ?? categories.FirstOrDefault();
            UpdateFamiliesMapping();
            SelectedFamilyMapping = SmartSuggestFamily(DisplayBlockName, FamiliesMapping) ?? FamiliesMapping.FirstOrDefault();
            UpdateTypeSymbolsMapping();
            SelectedTypeSymbolMapping = SmartSuggestTypeSymbol(DisplayBlockName, TypeSymbolsMapping) ?? TypeSymbolsMapping.FirstOrDefault();
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

        public string DisplayBlockName
        {
            get
            {
                if (string.IsNullOrEmpty(BlockName))
                    return BlockName;
                int prefixIndex = BlockName.LastIndexOf('.');
                return prefixIndex >= 0 ? BlockName.Substring(prefixIndex + 1) : BlockName;
            }
        }

        private ObservableCollection<Category> _categoriesMapping;
        public ObservableCollection<Category> CategoriesMapping
        {
            get => _categoriesMapping;
            set { _categoriesMapping = value; OnPropertyChanged(nameof(CategoriesMapping)); }
        }

        private Category _selectedCategoryMapping;
        public Category SelectedCategoryMapping
        {
            get => _selectedCategoryMapping;
            set
            {
                _selectedCategoryMapping = value;
                OnPropertyChanged(nameof(SelectedCategoryMapping));
                UpdateFamiliesMapping();
                UpdateTypeSymbolsMapping();
            }
        }

        private ObservableCollection<Family> _familiesMapping;
        public ObservableCollection<Family> FamiliesMapping
        {
            get => _familiesMapping;
            set { _familiesMapping = value; OnPropertyChanged(nameof(FamiliesMapping)); }
        }

        private Family _selectedFamilyMapping;
        public Family SelectedFamilyMapping
        {
            get => _selectedFamilyMapping;
            set
            {
                _selectedFamilyMapping = value;
                OnPropertyChanged(nameof(SelectedFamilyMapping));
                UpdateTypeSymbolsMapping();
            }
        }

        private ObservableCollection<FamilySymbol> _typeSymbolsMapping;
        public ObservableCollection<FamilySymbol> TypeSymbolsMapping
        {
            get => _typeSymbolsMapping;
            set { _typeSymbolsMapping = value; OnPropertyChanged(nameof(TypeSymbolsMapping)); }
        }

        private FamilySymbol _selectedTypeSymbolMapping;
        public FamilySymbol SelectedTypeSymbolMapping
        {
            get => _selectedTypeSymbolMapping;
            set { _selectedTypeSymbolMapping = value; OnPropertyChanged(nameof(SelectedTypeSymbolMapping)); }
        }
        private double _offset;
        public double Offset
        {
            get => _offset;
            set { _offset = value; OnPropertyChanged(nameof(Offset)); }
        }

        // -------- AI Suggest (auto) ----------
        private static IEnumerable<string> GetTokens(string name)
        {
            return name?.Split('_', '-', ' ', '.', '@')
                        .Select(s => s.ToLower())
                        .Where(s => s.Length > 2) ?? Enumerable.Empty<string>();
        }

        private Category SmartSuggestCategory(string blockName, IEnumerable<Category> categories)
        {
            var tokens = GetTokens(blockName).ToList();
            return categories.OrderByDescending(cat =>
                tokens.Count(token => cat.Name.ToLower().Contains(token)))
                .FirstOrDefault();
        }

        private Family SmartSuggestFamily(string blockName, IEnumerable<Family> families)
        {
            var tokens = GetTokens(blockName).ToList();
            return families.OrderByDescending(fam =>
                tokens.Count(token => fam.Name.ToLower().Contains(token)))
                .FirstOrDefault();
        }

        private FamilySymbol SmartSuggestTypeSymbol(string blockName, IEnumerable<FamilySymbol> typeSymbols)
        {
            var tokens = GetTokens(blockName).ToList();
            return typeSymbols.OrderByDescending(fs =>
                tokens.Count(token => fs.Name.ToLower().Contains(token)))
                .FirstOrDefault();
        }
        // -------- End AI Suggest -------------

        private void UpdateFamiliesMapping()
        {
            if (SelectedCategoryMapping != null && ThisDoc != null)
            {
                FamiliesMapping = new ObservableCollection<Family>(
                    new FilteredElementCollector(ThisDoc)
                        .OfClass(typeof(Family))
                        .Cast<Family>()
                        .Where(family => family.FamilyCategory.Id == SelectedCategoryMapping.Id)
                        .OrderBy(family => family.Name));
                SelectedFamilyMapping = FamiliesMapping.FirstOrDefault();
            }
        }

        private void UpdateTypeSymbolsMapping()
        {
            if (SelectedFamilyMapping != null && ThisDoc != null)
            {
                TypeSymbolsMapping = new ObservableCollection<FamilySymbol>(
                    new FilteredElementCollector(ThisDoc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .Where(fs => fs.Family.Id == SelectedFamilyMapping.Id)
                        .OrderBy(fs => fs.Name));
                SelectedTypeSymbolMapping = TypeSymbolsMapping.FirstOrDefault();
            }
        }

        protected void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class ImportInstanceSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is ImportInstance;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}
