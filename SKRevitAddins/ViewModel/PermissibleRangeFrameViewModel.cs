using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SKRevitAddins.Utils;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Document = Autodesk.Revit.DB.Document;

namespace SKRevitAddins.ViewModel
{
    public class PermissibleRangeFrameViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Application ThisApp;
        private Document ThisDoc;

        public PermissibleRangeFrameViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisApp = UiApp.Application;
            ThisDoc = UiDoc.Document;

            var refLinkCad = UiDoc.Selection.PickObject(ObjectType.Element, new ImportInstanceSelectionFilter(), "Select Link File");
            SelectedCadLink = ThisDoc.GetElement(refLinkCad) as ImportInstance;
            if (SelectedCadLink == null) return;
            AllLayers = CadUtils.GetAllLayer(SelectedCadLink);
            SelectedLayer = AllLayers[0];

            PileType = new FilteredElementCollector(ThisDoc).WhereElementIsElementType().OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .Where(e => e.Name.Contains("杭")).ToList();

            SelectedPileType = PileType[0] as FamilySymbol;
            AllLevel = new FilteredElementCollector(ThisDoc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Levels)
                .OfClass(typeof(Level)).Cast<Level>()
                .ToList();
            SelectedLevel = AllLevel[0];
            Offset = 0;
        }

        #region Properties
        /// <summary>
        /// Assign value of conditions
        /// </summary>

        public double X { get; set; }
        public double Y { get; set; }
        public double A { get; set; }
        public double B { get; set; }
        public double C { get; set; }
        public double D { get; set; }


        private ImportInstance _selectedCadLink;
        public ImportInstance SelectedCadLink
        {
            get { return _selectedCadLink; }
            set { _selectedCadLink = value; OnPropertyChanged(nameof(SelectedCadLink)); }
        }
        private List<string> _allLayers { get; set; }
        public List<string> AllLayers
        {
            get { return _allLayers; }
            set { _allLayers = value; OnPropertyChanged(nameof(AllLayers)); }
        }
        private string _selectedLayer { get; set; }
        public string SelectedLayer
        {
            get { return _selectedLayer; }
            set { _selectedLayer = value; OnPropertyChanged(nameof(SelectedLayer)); }
        }
        private IList<Element> _pileType { get; set; }
        public IList<Element> PileType
        {
            get { return _pileType; }
            set { _pileType = value; OnPropertyChanged(nameof(PileType)); }
        }
        private FamilySymbol _selectedPileType { get; set; }

        public FamilySymbol SelectedPileType
        {
            get { return _selectedPileType; }
            set { _selectedPileType = value; OnPropertyChanged(nameof(SelectedPileType)); }
        }
        private List<Level> _allLevel { get; set; }
        public List<Level> AllLevel
        {
            get { return _allLevel; }
            set { _allLevel = value; OnPropertyChanged(nameof(AllLevel)); }
        }
        private Level _selectedLevel { get; set; }
        public Level SelectedLevel
        {
            get { return _selectedLevel; }
            set { _selectedLevel = value; OnPropertyChanged(nameof(SelectedLevel)); }
        }
        private double _offset{ get; set; }
        public double Offset
        {
            get { return _offset; }
            set { _offset = value; OnPropertyChanged(nameof(Offset)); }
        }
        public Document Doc { get; set; }
        public class FrameObj : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            public FrameObj(Element frameObj)
            {
                FramingObj = frameObj;
                FramingGeometryObject = frameObj.get_Geometry(new Options());
                if (FramingGeometryObject == null) return;
                List<Solid> solids = ElementGeometryUtils.GetSolidsFromGeometry(FramingGeometryObject);
                FramingSolid = solids.UnionSolidList();
                FramingHeight = FramingSolid.GetSolidHeight();
            }
            public Element FramingObj { get; private set; }
            public GeometryElement FramingGeometryObject { get; private set; }
            public Solid FramingSolid { get; private set; }
            public double FramingHeight { get; private set; }

            protected void OnPropertyChanged(string propertyName)
            {
                if (PropertyChanged != null) // if there is any subscribers 
                    PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
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

        #endregion
    }
}
