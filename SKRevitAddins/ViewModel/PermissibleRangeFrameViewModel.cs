using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SKRevitAddins.Utils;

namespace SKRevitAddins.ViewModel
{
    public class PermissibleRangeFrameViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Document ThisDoc;

        public PermissibleRangeFrameViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisDoc = UiDoc.Document;

            X = 1.0 / 2.0;
            Y = 1.0 / 4.0;
            A = 750.0;
            B = 1.0 / 3.0;
            C = 2.0 / 3.0;

            PermissibleRange = true;
            PlaceSleeves = true;
            CreateErrorSchedules = true;

            DirectShapes = new List<DirectShape>();
            IntersectionData = new Dictionary<(ElementId MEPCurveId, ElementId FrameId), List<XYZ>>();
            ErrorMessages = new Dictionary<ElementId, HashSet<string>>();
            SleevePlacements = new Dictionary<ElementId, List<(XYZ, double)>>();

            List<Document> linkedDocs = new FilteredElementCollector(ThisDoc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Select(link => link.GetLinkDocument())
                .Where(linkedDoc => linkedDoc != null)
                .ToList();

            foreach (var linkedDoc in linkedDocs)
            {
                StructuralFramings = new FilteredElementCollector(linkedDoc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .ToList();
            }

            SleeveSymbol = new FilteredElementCollector(ThisDoc)
                .OfCategory(BuiltInCategory.OST_PipeAccessory)
                .OfClass(typeof(FamilySymbol))
                .WhereElementIsElementType()
                .Cast<FamilySymbol>()
                .FirstOrDefault(symbol => symbol.FamilyName == "スリーブ_SK");
            if (SleeveSymbol != null) return;
            TaskDialog.Show("Thông báo:", $"Không tìm thấy Family スリーブ_SK.");

            //MepCurves = new FilteredElementCollector(ThisDoc, ThisDoc.ActiveView.Id)
            //    .OfClass(typeof(MEPCurve))
            //    .Cast<MEPCurve>()
            //    .ToList();
        }

        #region Properties

        private double _x;
        public double X { get { return _x; } set { _x = value; OnPropertyChanged(nameof(X)); } }
        private double _y;
        public double Y { get { return _y; } set { _y = value; OnPropertyChanged(nameof(Y)); } }
        private double _a;
        public double A { get { return _a; } set { _a = value; OnPropertyChanged(nameof(A)); } }
        private double _b;
        public double B { get { return _b; } set { _b = value; OnPropertyChanged(nameof(B)); } }
        private double _c;
        public double C { get { return _c; } set { _c = value; OnPropertyChanged(nameof(C)); } }

        private FamilySymbol _sleeveSymbol;

        public FamilySymbol SleeveSymbol
        {
            get { return _sleeveSymbol; }
            set { _sleeveSymbol = value; OnPropertyChanged(nameof(SleeveSymbol)); }
        }

        private List<Element> _structuralFramings;
        public List<Element> StructuralFramings
        {
            get { return _structuralFramings; }
            set { _structuralFramings = value; OnPropertyChanged(nameof(StructuralFramings)); }
        }

        private List<DirectShape> _directShapes;
        public List<DirectShape> DirectShapes
        {
            get { return _directShapes; }
            set { _directShapes = value; OnPropertyChanged(nameof(DirectShapes)); }
        }
        private List<MEPCurve> _mepCurves;
        public List<MEPCurve> MepCurves
        {
            get { return _mepCurves; }
            set { _mepCurves = value; OnPropertyChanged(nameof(MepCurves)); }
        }
        private Dictionary<(ElementId MEPCurveId, ElementId FrameId), List<XYZ>> _intersectionData;
        public Dictionary<(ElementId MEPCurveId, ElementId FrameId), List<XYZ>> IntersectionData
        {
            get { return _intersectionData; }
            set { _intersectionData = value; OnPropertyChanged(nameof(IntersectionData)); }
        }

        private Dictionary<ElementId, HashSet<string>> _errorMessages;
        public Dictionary<ElementId, HashSet<string>> ErrorMessages
        {
            get { return _errorMessages; }
            set { _errorMessages = value; OnPropertyChanged(nameof(ErrorMessages)); }
        }

        private Dictionary<ElementId, List<(XYZ, double)>> _sleevePlacements;
        public Dictionary<ElementId, List<(XYZ, double)>> SleevePlacements
        {
            get { return _sleevePlacements; }
            set { _sleevePlacements = value; OnPropertyChanged(nameof(SleevePlacements)); }
        }

        private bool _permissibleRange;
        public bool PermissibleRange
        {
            get { return _permissibleRange; }
            set { _permissibleRange = value; OnPropertyChanged(nameof(PermissibleRange)); }
        }

        private bool _placeSleeves;
        public bool PlaceSleeves
        {
            get { return _placeSleeves; }
            set { _placeSleeves = value; OnPropertyChanged(nameof(PlaceSleeves)); }
        }private bool _createErrorSchedules;
        public bool CreateErrorSchedules
        {
            get { return _createErrorSchedules; }
            set { _createErrorSchedules = value; OnPropertyChanged(nameof(CreateErrorSchedules)); }
        }

        public class FrameObj : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            public FrameObj(Element frameObj)
            {
                FramingObj = frameObj;
                FramingGeometryObj = frameObj.get_Geometry(new Options());
                if (FramingGeometryObj == null) return;
                List<Solid> solids = ElementGeometryUtils.GetSolidsFromGeometry(FramingGeometryObj);
                FramingSolid = solids.UnionSolidList();
                FramingHeight = FramingSolid.GetSolidHeight();
            }

            private Element _framingObj;
            public Element FramingObj 
            { 
                get { return _framingObj; }
                set { _framingObj = value; OnPropertyChanged(nameof(FramingObj)); }
            }
            private GeometryElement _framingGeometryObj;

            public GeometryElement FramingGeometryObj
            {
                get { return _framingGeometryObj; }
                set { _framingGeometryObj = value; OnPropertyChanged(nameof(FramingGeometryObj)); }
            }

            private Solid _framingSolid;

            public Solid FramingSolid
            {
                get { return _framingSolid; }
                set { _framingSolid = value; OnPropertyChanged(nameof(FramingSolid)); }
            }

            private double _framingHeight;
            public double FramingHeight
            {
                get { return _framingHeight; }
                set { _framingHeight = value; OnPropertyChanged(nameof(FramingHeight)); }
            }

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
