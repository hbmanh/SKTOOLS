using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using DocumentFormat.OpenXml.Drawing;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace SKToolsAddins.ViewModel
{
    public class PlaceElementsFromBlocksCadViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Document ThisDoc;

        public PlaceElementsFromBlocksCadViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisDoc = UiDoc.Document;

            // Chọn file CAD link
            var refLinkCad = UiDoc.Selection.PickObject(ObjectType.Element, new ImportInstanceSelectionFilter(), "Select Link File");
            var selectedCadLink = ThisDoc.GetElement(refLinkCad) as ImportInstance;
            if (selectedCadLink == null)
            {
                TaskDialog.Show("Error", "No valid CAD link selected.");
                return;
            }

            Categories = new ObservableCollection<Category>(GetCategories(ThisDoc));
            SelectedCategory = Categories.FirstOrDefault();

            UpdateFamilies();
            UpdateTypeSymbols();

            // Lấy danh sách block từ file CAD link
            var blocks = GetBlockNamesFromCadLink(selectedCadLink);
            foreach (var block in blocks)
            {
                BlockMappings.Add(new BlockMapping(block, ThisDoc, Categories, Families, TypeSymbols));
            }
        }

        #region Properties

        public ObservableCollection<BlockMapping> BlockMappings { get; set; } = new ObservableCollection<BlockMapping>();

        public ObservableCollection<Category> Categories { get; set; }

        private Category _selectedCategory;
        public Category SelectedCategory
        {
            get { return _selectedCategory; }
            set
            {
                _selectedCategory = value;
                OnPropertyChanged(nameof(SelectedCategory));
                UpdateFamilies();
                UpdateTypeSymbols();
            }
        }

        public ObservableCollection<Family> Families { get; set; } = new ObservableCollection<Family>();

        private Family _selectedFamily;
        public Family SelectedFamily
        {
            get { return _selectedFamily; }
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
            get { return _selectedTypeSymbol; }
            set
            {
                _selectedTypeSymbol = value;
                OnPropertyChanged(nameof(SelectedTypeSymbol));
            }
        }

        #endregion

        #region Methods

        private void UpdateFamilies()
        {
            if (SelectedCategory != null)
            {
                Families = new ObservableCollection<Family>(
                    new FilteredElementCollector(ThisDoc)
                        .OfClass(typeof(Family))
                        .Cast<Family>()
                        .Where(family => family.FamilyCategory.Id == SelectedCategory.Id)
                        .OrderBy(family => family.Name)
                        .ToList());
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
                        .OrderBy(fs => fs.Name)
                        .ToList());
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

        private static List<GeometryInstance> GetBlockNamesFromCadLink(ImportInstance cadLink)
        {
            var blocks = new List<GeometryInstance>();
            GeometryElement geoElement = cadLink.get_Geometry(new Options());

            foreach (GeometryObject geoObject in geoElement)
            {
                if (geoObject is GeometryInstance instance)
                {
                    foreach (GeometryObject instObj in instance.SymbolGeometry)
                    {
                        if (instObj is GeometryInstance blockInstance && blocks.All(b => b.Symbol.Name != blockInstance.Symbol.Name))
                        {
                            blocks.Add(blockInstance);
                        }
                    }
                }
            }

            return blocks;
        }

        #endregion
    }

    public class BlockMapping : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public Document ThisDoc { get; set; }

        public BlockMapping(GeometryInstance block, Document doc, ObservableCollection<Category> categories, ObservableCollection<Family> families, ObservableCollection<FamilySymbol> typeSymbols)
        {
            ThisDoc = doc;
            Block = block;
            BlockName = block.Symbol.Name;
            CategoriesMapping = categories;
            FamiliesMapping = families;
            TypeSymbolsMapping = typeSymbols;
            SelectedCategoryMapping = categories.FirstOrDefault();
            Offset = 2600;
            UpdateFamiliesMapping();
            UpdateTypeSymbolsMapping();
        }

        private GeometryInstance _block;
        public GeometryInstance Block
        {
            get { return _block; }
            set
            {
                _block = value;
                OnPropertyChanged(nameof(Block));
            }
        }

        private string _blockName;
        public string BlockName
        {
            get { return _blockName; }
            set
            {
                _blockName = value;
                OnPropertyChanged(nameof(BlockName));
                OnPropertyChanged(nameof(DisplayBlockName));
            }
        }

        // New property to display the block name without the prefix
        public string DisplayBlockName
        {
            get
            {
                if (string.IsNullOrEmpty(BlockName))
                    return BlockName;

                // Remove the prefix (e.g., "Debug.dwg.")
                int prefixIndex = BlockName.LastIndexOf('.');
                return prefixIndex >= 0 ? BlockName.Substring(prefixIndex + 1) : BlockName;
            }
        }

        private ObservableCollection<Category> _categoriesMapping;
        public ObservableCollection<Category> CategoriesMapping
        {
            get { return _categoriesMapping; }
            set
            {
                _categoriesMapping = value;
                OnPropertyChanged(nameof(CategoriesMapping));
            }
        }

        private Category _selectedCategoryMapping;
        public Category SelectedCategoryMapping
        {
            get { return _selectedCategoryMapping; }
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
            get { return _familiesMapping; }
            set
            {
                _familiesMapping = value;
                OnPropertyChanged(nameof(FamiliesMapping));
            }
        }

        private Family _selectedFamilyMapping;
        public Family SelectedFamilyMapping
        {
            get { return _selectedFamilyMapping; }
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
            get { return _typeSymbolsMapping; }
            set
            {
                _typeSymbolsMapping = value;
                OnPropertyChanged(nameof(TypeSymbolsMapping));
            }
        }

        private FamilySymbol _selectedTypeSymbolMapping;
        public FamilySymbol SelectedTypeSymbolMapping
        {
            get { return _selectedTypeSymbolMapping; }
            set
            {
                _selectedTypeSymbolMapping = value;
                OnPropertyChanged(nameof(SelectedTypeSymbolMapping));
            }
        }
        private double _offset;
        public double Offset
        {
            get { return _offset; }
            set
            {
                _offset = value;
                OnPropertyChanged(nameof(Offset));
            }
        }

        private void UpdateFamiliesMapping()
        {
            if (SelectedCategoryMapping != null && ThisDoc != null)
            {
                FamiliesMapping = new ObservableCollection<Family>(
                    new FilteredElementCollector(ThisDoc)
                        .OfClass(typeof(Family))
                        .Cast<Family>()
                        .Where(family => family.FamilyCategory.Id == SelectedCategoryMapping.Id)
                        .OrderBy(family => family.Name)
                        .ToList());
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
                        .OrderBy(fs => fs.Name)
                        .ToList());
                SelectedTypeSymbolMapping = TypeSymbolsMapping.FirstOrDefault();
            }
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ImportInstanceSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is ImportInstance;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
