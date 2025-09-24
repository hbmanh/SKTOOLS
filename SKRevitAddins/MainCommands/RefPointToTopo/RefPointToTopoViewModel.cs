using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace SKRevitAddins.RefPointToTopo
{
    public class RelayCommand : ICommand
    {
        readonly Action<object> _run; readonly Func<object, bool> _can;
        public RelayCommand(Action<object> run, Func<object, bool> can = null) { _run = run; _can = can; }
        public bool CanExecute(object p) => _can?.Invoke(p) ?? true;
        public void Execute(object p) => _run(p);
        public event EventHandler CanExecuteChanged { add => CommandManager.RequerySuggested += value; remove => CommandManager.RequerySuggested -= value; }
    }

    public class RefPointToTopoViewModel : INotifyPropertyChanged
    {
        readonly UIDocument _uiDoc;
        readonly RefPointToTopoHandler _handler = new();
        readonly ExternalEvent _evt;

        public RefPointToTopoViewModel(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;
            _evt = ExternalEvent.Create(_handler);

            PickTopoCmd = new RelayCommand(_ => PickToposolid());
            RunCmd = new RelayCommand(_ => Start(), _ => TargetTopoId != null && !IsBusy);
            CancelCmd = new RelayCommand(_ => Cancel(), _ => IsBusy);
        }

        // ===== Bindable =====
        ElementId _targetTopoId;
        public ElementId TargetTopoId { get => _targetTopoId; set { _targetTopoId = value; OnChanged(); OnChanged(nameof(TargetTopoText)); CommandManager.InvalidateRequerySuggested(); } }
        public string TargetTopoText => TargetTopoId?.IntegerValue > 0 ? TargetTopoId.IntegerValue.ToString() : "(chưa chọn)";

        string _refFam = "RefPoint";
        public string RefPointFamilyName { get => _refFam; set { _refFam = value; OnChanged(); } }

        // —— Đơn vị nhập: mm ——
        double _gridMm = 2000.0, _edgeMm = 1000.0, _snapMm = 50.0;
        public double GridSpacingMillimeters { get => _gridMm; set { _gridMm = value; OnChanged(); } }
        public double EdgeSpacingMillimeters { get => _edgeMm; set { _edgeMm = value; OnChanged(); } }
        public double SnapDistanceMillimeters { get => _snapMm; set { _snapMm = value; OnChanged(); } }

        int _k = 6; public int IDW_K { get => _k; set { _k = value; OnChanged(); } }
        double _pow = 2.0; public double IDW_Power { get => _pow; set { _pow = value; OnChanged(); } }
        int _maxPts = 18000; public int MaxPoints { get => _maxPts; set { _maxPts = value; OnChanged(); } }

        bool _adaptive = true; public bool UseAdaptiveSampling { get => _adaptive; set { _adaptive = value; OnChanged(); } }
        double _grad = 0.30; public double GradThreshold { get => _grad; set { _grad = value; OnChanged(); } }
        double _refine = 0.5; public double RefineFactor { get => _refine; set { _refine = value; OnChanged(); } }
        int _maxRef = 6000; public int MaxRefinePoints { get => _maxRef; set { _maxRef = value; OnChanged(); } }
        bool _parallel = true; public bool UseParallelIDW { get => _parallel; set { _parallel = value; OnChanged(); } }

        bool _busy; public bool IsBusy { get => _busy; set { _busy = value; OnChanged(); CommandManager.InvalidateRequerySuggested(); } }
        double _progress; public double Progress { get => _progress; set { _progress = value; OnChanged(); } }

        public ICommand PickTopoCmd { get; }
        public ICommand RunCmd { get; }
        public ICommand CancelCmd { get; }

        void PickToposolid()
        {
            try
            {
                var r = _uiDoc.Selection.PickObject(ObjectType.Element, new TopoFilter(), "Chọn Toposolid muốn chỉnh sửa");
                TargetTopoId = r.ElementId;
            }
            catch { /* user cancel */ }
        }

        void Start()
        {
            // MM -> FEET
            double gridFt = UnitUtils.ConvertToInternalUnits(GridSpacingMillimeters, UnitTypeId.Millimeters);
            double edgeFt = UnitUtils.ConvertToInternalUnits(EdgeSpacingMillimeters, UnitTypeId.Millimeters);
            double snapFt = UnitUtils.ConvertToInternalUnits(SnapDistanceMillimeters, UnitTypeId.Millimeters);

            _handler.TargetToposolidId = TargetTopoId;
            _handler.RefPointFamilyName = RefPointFamilyName;
            _handler.GridSpacingFt = gridFt;
            _handler.EdgeSpacingFt = edgeFt;
            _handler.SnapDistanceFt = snapFt;
            _handler.IDW_K = IDW_K;
            _handler.IDW_Power = IDW_Power;
            _handler.MaxPoints = MaxPoints;
            _handler.UseAdaptiveSampling = UseAdaptiveSampling;
            _handler.GradThreshold = GradThreshold;
            _handler.RefineFactor = RefineFactor;
            _handler.MaxRefinePoints = MaxRefinePoints;
            _handler.UseParallelIDW = UseParallelIDW;

            _handler.BusySetter = v => IsBusy = v;
            _handler.ProgressReporter = (c, t) => Progress = t == 0 ? 0 : (double)c / t;
            _handler.IsCancelled = false;

            IsBusy = true; Progress = 0;
            _evt.Raise();
        }

        void Cancel() => _handler.IsCancelled = true;

        class TopoFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem?.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_Toposolid;
            public bool AllowReference(Reference reference, XYZ position) => true;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        void OnChanged([CallerMemberName] string n = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }
}
