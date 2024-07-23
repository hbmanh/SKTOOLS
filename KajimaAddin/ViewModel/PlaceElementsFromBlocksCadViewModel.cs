using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using static SKToolsAddins.ViewModel.AutoCreatePileFromCadViewModel;
using static SKToolsAddins.ViewModel.BlockMapping;

namespace SKToolsAddins.ViewModel
{
    public class PlaceElementsFromBlocksCadViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Application ThisApp;
        private Document ThisDoc;

        public PlaceElementsFromBlocksCadViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisApp = UiApp.Application;
            ThisDoc = UiDoc.Document;

            // Chọn file CAD link
            var refLinkCad = UiDoc.Selection.PickObject(ObjectType.Element, new ImportInstanceSelectionFilter(), "Select Link File");
            var selectedCadLink = ThisDoc.GetElement(refLinkCad) as ImportInstance;
            if (selectedCadLink == null)
            {
                TaskDialog.Show("Error", "No valid CAD link selected.");
                return;
            }
            Categories = new ObservableCollection<Category>(new FilteredElementCollector(ThisDoc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Select(family => family.FamilyCategory)
                .Where(c => c != null && c.CategoryType == CategoryType.Model)
                .GroupBy(c => c.Id)
                .Select(g => g.First())
                .OrderBy(c => c.Name)
                .ToList());
            SelectedCategory = Categories.FirstOrDefault();

            // Lấy danh sách block từ file CAD link
            var blockNames = GetBlockNamesFromCadLink(selectedCadLink);
            foreach (var blockName in blockNames)
            {
                BlockMappings.Add(new BlockMapping
                {
                    BlockName = blockName,
                    CategoriesMapping = Categories,
                    FamiliesMapping = Families,
                    TypeSymbolsMapping = TypeSymbols
                    
                });
            }

        }

        #region Properties

        private ObservableCollection<BlockMapping> _blockMappings = new ObservableCollection<BlockMapping>();
        public ObservableCollection<BlockMapping> BlockMappings
        {
            get { return _blockMappings; }
            set
            {
                _blockMappings = value;
                OnPropertyChanged(nameof(BlockMappings));
            }
        }

        private ObservableCollection<Category> _categories;
        public ObservableCollection<Category> Categories
        {
            get { return _categories; }
            set
            {
                _categories = value;
                OnPropertyChanged(nameof(Categories));
            }
        }

        private Category _selectedCategory;
        public Category SelectedCategory
        {
            get { return _selectedCategory; }
            set
            {
                _selectedCategory = value;
                OnPropertyChanged(nameof(SelectedCategory));
                UpdateFamilySymbols();
                UpdateTypeSymbols();
            }
        }
        private ObservableCollection<Family> _families;
        public ObservableCollection<Family> Families
        {
            get { return _families; }
            set
            {
                _families = value;
                OnPropertyChanged(nameof(Families));
            }
        }
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

        private ObservableCollection<FamilySymbol> _typeSymbols;
        public ObservableCollection<FamilySymbol> TypeSymbols
        {
            get { return _typeSymbols; }
            set
            {
                _typeSymbols = value;
                OnPropertyChanged(nameof(TypeSymbols));
            }
        }

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
        public void UpdateFamilySymbols()
        {
            // Cập nhật danh sách Family Symbol dựa trên Category đã chọn
            if (SelectedCategory != null)
            {
                var familySymbols = new FilteredElementCollector(ThisDoc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .Where(family => family.FamilyCategory.Id == SelectedCategory.Id)
                    .OrderBy(family => family.Name)
                    .ToList();
                Families = new ObservableCollection<Family>(familySymbols);
                SelectedFamily = Families[0];
            }
        }

        public void UpdateTypeSymbols()
        {
            // Cập nhật danh sách Type Symbol dựa trên Family Symbol đã chọn
            if (SelectedFamily != null)
            {
                var typeSymbols = new FilteredElementCollector(ThisDoc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Where(fs => fs.Family.Id == SelectedFamily.Id)
                    .OrderBy(fs => fs.Name)
                    .ToList();
            
                TypeSymbols = new ObservableCollection<FamilySymbol>(typeSymbols);
                SelectedTypeSymbol = TypeSymbols[0];
            }
        }

        private List<string> GetBlockNamesFromCadLink(ImportInstance cadLink)
        {
            var blockNames = new List<string>();
            GeometryElement geoElement = cadLink.get_Geometry(new Options());

            foreach (GeometryObject geoObject in geoElement)
            {
                GeometryInstance instance = geoObject as GeometryInstance;
                if (instance != null)
                {
                    foreach (GeometryObject instObj in instance.SymbolGeometry)
                    {
                        if (instObj is GeometryInstance blockInstance)
                        {
                            blockNames.Add(blockInstance.Symbol.Name);
                        }
                    }
                }
            }

            return blockNames.Distinct().ToList();
        }

        #endregion
    }

    public class BlockMapping : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public Document ThisDoc { get; set; } // Thêm thuộc tính này để lưu trữ ThisDoc

        private string _blockName;
        public string BlockName
        {
            get { return _blockName; }
            set
            {
                _blockName = value;
                OnPropertyChanged(nameof(BlockName));
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


        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
