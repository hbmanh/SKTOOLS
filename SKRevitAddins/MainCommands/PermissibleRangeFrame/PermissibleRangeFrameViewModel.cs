using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SKRevitAddins.Utils;
using SKRevitAddins.PermissibleRangeFrame;   // thêm cuối danh sách using


namespace SKRevitAddins.PermissibleRangeFrame
{
    public class PermissibleRangeFrameViewModel : ViewModelBase
    {
        public UIApplication UiApp { get; }
        private UIDocument UiDoc { get; }
        private Document ThisDoc { get; }

        public PermissibleRangeFrameViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = uiApp.ActiveUIDocument;
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

            // Lấy các dầm từ tài liệu liên kết (RevitLinkInstance)
            List<Document> linkedDocs = new FilteredElementCollector(ThisDoc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Select(link => link.GetLinkDocument())
                .Where(linkedDoc => linkedDoc != null)
                .ToList();

            StructuralFramings = linkedDocs.SelectMany(linkedDoc => new FilteredElementCollector(linkedDoc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .ToList())
                .ToList();

            // Lấy Sleeve Symbol theo FamilyName "スリーブ_SK"
            SleeveSymbol = new FilteredElementCollector(ThisDoc)
                .OfCategory(BuiltInCategory.OST_PipeAccessory)
                .OfClass(typeof(FamilySymbol))
                .WhereElementIsElementType()
                .Cast<FamilySymbol>()
                .FirstOrDefault(symbol => symbol.FamilyName == "スリーブ_SK");
            if (SleeveSymbol == null)
                TaskDialog.Show("Thông báo:", "Không tìm thấy Family スリーブ_SK.");
        }

        #region Properties

        private double _x;
        public double X { get => _x; set { _x = value; OnPropertyChanged(nameof(X)); } }

        private double _y;
        public double Y { get => _y; set { _y = value; OnPropertyChanged(nameof(Y)); } }

        private double _a;
        public double A { get => _a; set { _a = value; OnPropertyChanged(nameof(A)); } }

        private double _b;
        public double B { get => _b; set { _b = value; OnPropertyChanged(nameof(B)); } }

        private double _c;
        public double C { get => _c; set { _c = value; OnPropertyChanged(nameof(C)); } }

        private FamilySymbol _sleeveSymbol;
        public FamilySymbol SleeveSymbol
        {
            get => _sleeveSymbol;
            set { _sleeveSymbol = value; OnPropertyChanged(nameof(SleeveSymbol)); OnPropertyChanged(nameof(CanCreate)); }
        }

        private List<Element> _structuralFramings;
        public List<Element> StructuralFramings { get => _structuralFramings; set { _structuralFramings = value; OnPropertyChanged(nameof(StructuralFramings)); } }

        private List<DirectShape> _directShapes;
        public List<DirectShape> DirectShapes { get => _directShapes; set { _directShapes = value; OnPropertyChanged(nameof(DirectShapes)); } }

        private List<MEPCurve> _mepCurves;
        public List<MEPCurve> MepCurves { get => _mepCurves; set { _mepCurves = value; OnPropertyChanged(nameof(MepCurves)); } }

        private Dictionary<(ElementId MEPCurveId, ElementId FrameId), List<XYZ>> _intersectionData;
        public Dictionary<(ElementId MEPCurveId, ElementId FrameId), List<XYZ>> IntersectionData { get => _intersectionData; set { _intersectionData = value; OnPropertyChanged(nameof(IntersectionData)); } }

        private Dictionary<ElementId, HashSet<string>> _errorMessages;
        public Dictionary<ElementId, HashSet<string>> ErrorMessages { get => _errorMessages; set { _errorMessages = value; OnPropertyChanged(nameof(ErrorMessages)); } }

        private Dictionary<ElementId, List<(XYZ, double)>> _sleevePlacements;
        public Dictionary<ElementId, List<(XYZ, double)>> SleevePlacements { get => _sleevePlacements; set { _sleevePlacements = value; OnPropertyChanged(nameof(SleevePlacements)); } }

        private bool _permissibleRange;
        public bool PermissibleRange { get => _permissibleRange; set { _permissibleRange = value; OnPropertyChanged(nameof(PermissibleRange)); } }

        private bool _placeSleeves;
        public bool PlaceSleeves { get => _placeSleeves; set { _placeSleeves = value; OnPropertyChanged(nameof(PlaceSleeves)); } }

        private bool _createErrorSchedules;
        public bool CreateErrorSchedules { get => _createErrorSchedules; set { _createErrorSchedules = value; OnPropertyChanged(nameof(CreateErrorSchedules)); } }

        // Quan trọng: property này dùng để disable nút OK nếu thiếu Family Symbol
        public bool CanCreate => SleeveSymbol != null;

        #endregion

        public class FrameObj : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            public FrameObj(Element frameObj)
            {
                FramingObj = frameObj;
                FramingGeometryObj = frameObj.get_Geometry(new Options());
                if (FramingGeometryObj == null) return;
                var solids = ElementGeometryUtils.GetSolidsFromGeometry(FramingGeometryObj);
                FramingSolid = solids.UnionSolidList();
                FramingHeight = FramingSolid.GetSolidHeight();
            }

            private Element _framingObj;
            public Element FramingObj { get => _framingObj; set { _framingObj = value; OnPropertyChanged(nameof(FramingObj)); } }

            private GeometryElement _framingGeometryObj;
            public GeometryElement FramingGeometryObj { get => _framingGeometryObj; set { _framingGeometryObj = value; OnPropertyChanged(nameof(FramingGeometryObj)); } }

            private Solid _framingSolid;
            public Solid FramingSolid { get => _framingSolid; set { _framingSolid = value; OnPropertyChanged(nameof(FramingSolid)); } }

            private double _framingHeight;
            public double FramingHeight { get => _framingHeight; set { _framingHeight = value; OnPropertyChanged(nameof(FramingHeight)); } }

            protected void OnPropertyChanged(string propertyName) =>
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public class ImportInstanceSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is ImportInstance;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }
    }
}
