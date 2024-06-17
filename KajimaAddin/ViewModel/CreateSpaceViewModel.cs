using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB.Mechanical;
using Application = Autodesk.Revit.ApplicationServices.Application;
using View = Autodesk.Revit.DB.View;

namespace SKToolsAddins.ViewModel
{
    public class CreateSpaceViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Application ThisApp;
        private Document ThisDoc;

        public CreateSpaceViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisApp = UiApp.Application;
            ThisDoc = UiDoc.Document;

            Phases = new ObservableCollection<Phase>(new FilteredElementCollector(ThisDoc)
                .OfCategory(BuiltInCategory.OST_Phases)
                .OfClass(typeof(Phase))
                .Cast<Phase>()
                .ToList());

            SelPhase = Phases.FirstOrDefault(p => p.Name.Equals(ThisDoc.ActiveView.get_Parameter(BuiltInParameter.VIEW_PHASE).AsValueString()));

            Views = new ObservableCollection<View>(new FilteredElementCollector(ThisDoc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.ViewType != ViewType.ThreeD)
                .OrderBy(a => a.Name)
                .ToList());

            SelectedViews = new ObservableCollection<View>();

            ListExistSpaces = new ObservableCollection<Space>(new FilteredElementCollector(ThisDoc)
                .OfCategory(BuiltInCategory.OST_MEPSpaces)
                .WhereElementIsNotElementType()
                .Cast<Space>()
                .ToList());

            ListTagTypeSpace = new ObservableCollection<FamilySymbol>(new FilteredElementCollector(ThisDoc)
                .WhereElementIsElementType()
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(f => f.Category.Name.Equals("Space Tags"))
                .ToList());

            SelTagTypeSpace = ListTagTypeSpace[0];
        }

        #region Properties

        private ObservableCollection<View> _views;

        public ObservableCollection<View> Views
        {
            get { return _views; }
            set
            {
                _views = value;
                OnPropertyChanged(nameof(Views));
            }
        }

        private ObservableCollection<View> _selectedViews;

        public ObservableCollection<View> SelectedViews
        {
            get { return _selectedViews; }
            set
            {
                _selectedViews = value;
                OnPropertyChanged(nameof(SelectedViews));
            }
        }

        private ObservableCollection<Space> _listExistSpaces;

        public ObservableCollection<Space> ListExistSpaces
        {
            get { return _listExistSpaces; }
            set
            {
                _listExistSpaces = value;
                OnPropertyChanged(nameof(ListExistSpaces));
            }
        }

        private ObservableCollection<FamilySymbol> _listTagTypeSpace;

        public ObservableCollection<FamilySymbol> ListTagTypeSpace
        {
            get { return _listTagTypeSpace; }
            set
            {
                _listTagTypeSpace = value;
                OnPropertyChanged(nameof(ListTagTypeSpace));
            }
        }

        private FamilySymbol _selTagTypeSpace;

        public FamilySymbol SelTagTypeSpace
        {
            get { return _selTagTypeSpace; }
            set
            {
                _selTagTypeSpace = value;
                OnPropertyChanged(nameof(SelTagTypeSpace));
            }
        }

        private int _setSpaceOffet;

        public int SetSpaceOffet
        {
            get { return _setSpaceOffet; }
            set
            {
                _setSpaceOffet = value;
                OnPropertyChanged(nameof(SetSpaceOffet));
            }
        }

        private ObservableCollection<Phase> _phases;

        public ObservableCollection<Phase> Phases
        {
            get { return _phases; }
            set
            {
                _phases = value;
                OnPropertyChanged(nameof(Phases));
            }
        }

        private Phase _selPhase;

        public Phase SelPhase
        {
            get { return _selPhase; }
            set
            {
                _selPhase = value;
                OnPropertyChanged(nameof(SelPhase));
            }
        }

        public bool _tagPlacementBOX;

        public bool TagPlacementBOX
        {
            get { return _tagPlacementBOX; }
            set
            {
                _tagPlacementBOX = value;
                OnPropertyChanged(nameof(TagPlacementBOX));
            }
        }

        private bool _nameNumberBOX;

        public bool NameNumberBOX
        {
            get { return _nameNumberBOX; }
            set
            {
                _nameNumberBOX = value;
                OnPropertyChanged(nameof(NameNumberBOX));
            }
        }

        private bool _spaceOffsetBOX;

        public bool SpaceOffsetBOX
        {
            get { return _spaceOffsetBOX; }
            set
            {
                _spaceOffsetBOX = value;
                OnPropertyChanged(nameof(SpaceOffsetBOX));
            }
        }

        #endregion
    }
}
