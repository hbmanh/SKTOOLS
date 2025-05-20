using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace SKRevitAddins.CopySetOfFilterFromViewTemp
{
    public class CopySetOfFilterFromViewTempViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Application ThisApp;
        private Document ThisDoc;

        public CopySetOfFilterFromViewTempViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisApp = UiApp.Application;
            ThisDoc = UiDoc.Document;

            ViewTypes = new ObservableCollection<ViewType>(
                new FilteredElementCollector(ThisDoc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => v.IsTemplate && v.ViewType != ViewType.Schedule)
                    .Select(v => v.ViewType)
                    .Distinct()
                    .OrderBy(v => v.ToString())
                    .ToList());
            SelViewType = ViewTypes.FirstOrDefault();

            ViewsTypeTarget = new ObservableCollection<ViewType>(
                new FilteredElementCollector(ThisDoc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => v.ViewType != ViewType.Schedule)
                    .Select(v => v.ViewType)
                    .Distinct()
                    .OrderBy(v => v.ToString())
                    .ToList());
            SelViewTypeTarget = ViewsTypeTarget.FirstOrDefault();

            UpdateViewTemplate();
            SelViewTemplate = ViewTemplates.FirstOrDefault();

            UpdateFilters();
            SelFilter = new ObservableCollection<FilterObj>();

            UpdateViewTarget();
            SelViewTarget = new ObservableCollection<View>();
        }

        #region Properties

        private ObservableCollection<ViewType> _viewTypes;
        public ObservableCollection<ViewType> ViewTypes
        {
            get => _viewTypes;
            set { _viewTypes = value; OnPropertyChanged(nameof(ViewTypes)); }
        }

        private ViewType _selViewType;
        public ViewType SelViewType
        {
            get => _selViewType;
            set
            {
                _selViewType = value;
                OnPropertyChanged(nameof(SelViewType));
                UpdateViewTemplate();
            }
        }

        private ObservableCollection<ViewType> _viewsTypeTarget;
        public ObservableCollection<ViewType> ViewsTypeTarget
        {
            get => _viewsTypeTarget;
            set { _viewsTypeTarget = value; OnPropertyChanged(nameof(ViewsTypeTarget)); }
        }

        private ObservableCollection<View> _viewTemplates;
        public ObservableCollection<View> ViewTemplates
        {
            get => _viewTemplates;
            set { _viewTemplates = value; OnPropertyChanged(nameof(ViewTemplates)); }
        }

        private View _selViewTemplate;
        public View SelViewTemplate
        {
            get => _selViewTemplate;
            set
            {
                _selViewTemplate = value;
                OnPropertyChanged(nameof(SelViewTemplate));
                UpdateFilters();
            }
        }

        private ObservableCollection<FilterObj> _filters;
        public ObservableCollection<FilterObj> Filters
        {
            get => _filters;
            set { _filters = value; OnPropertyChanged(nameof(Filters)); }
        }

        private ObservableCollection<FilterObj> _selFilter;
        public ObservableCollection<FilterObj> SelFilter
        {
            get => _selFilter;
            set { _selFilter = value; OnPropertyChanged(nameof(SelFilter)); }
        }

        private ObservableCollection<View> _viewTargets;
        public ObservableCollection<View> ViewTargets
        {
            get => _viewTargets;
            set { _viewTargets = value; OnPropertyChanged(nameof(ViewTargets)); }
        }

        private ObservableCollection<View> _selViewTarget;
        public ObservableCollection<View> SelViewTarget
        {
            get => _selViewTarget;
            set { _selViewTarget = value; OnPropertyChanged(nameof(SelViewTarget)); }
        }

        private ViewType _selViewTypeTarget;
        public ViewType SelViewTypeTarget
        {
            get => _selViewTypeTarget;
            set
            {
                _selViewTypeTarget = value;
                OnPropertyChanged(nameof(SelViewTypeTarget));
                UpdateViewTarget();
            }
        }

        // --- Option Checkboxes ---
        private bool _allCopyBOX;
        public bool AllCopyBOX
        {
            get => _allCopyBOX;
            set
            {
                _allCopyBOX = value;
                OnPropertyChanged(nameof(AllCopyBOX));
                UpdateOptionState();
            }
        }

        private bool _patternCopyBOX;
        public bool PatternCopyBOX
        {
            get => _patternCopyBOX;
            set
            {
                _patternCopyBOX = value;
                if (_patternCopyBOX) AllCopyBOX = false;
                OnPropertyChanged(nameof(PatternCopyBOX));
                UpdateOptionState();
            }
        }

        private bool _cutSetCopyBOX;
        public bool CutSetCopyBOX
        {
            get => _cutSetCopyBOX;
            set
            {
                _cutSetCopyBOX = value;
                if (_cutSetCopyBOX) AllCopyBOX = false;
                OnPropertyChanged(nameof(CutSetCopyBOX));
                UpdateOptionState();
            }
        }

        private bool _patternCopyEnabled = true;
        public bool PatternCopyEnabled
        {
            get => _patternCopyEnabled;
            set { _patternCopyEnabled = value; OnPropertyChanged(nameof(PatternCopyEnabled)); }
        }

        private bool _cutSetCopyEnabled = true;
        public bool CutSetCopyEnabled
        {
            get => _cutSetCopyEnabled;
            set { _cutSetCopyEnabled = value; OnPropertyChanged(nameof(CutSetCopyEnabled)); }
        }

        public class FilterObj : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            public FilterObj(Document doc, ElementId filterElementId)
            {
                FilterId = filterElementId;
                FilterName = doc.GetElement(filterElementId).Name;
            }

            private ElementId _filterId;
            public ElementId FilterId
            {
                get => _filterId;
                set { _filterId = value; OnPropertyChanged(nameof(FilterId)); }
            }

            private string _filterName;
            public string FilterName
            {
                get => _filterName;
                set { _filterName = value; OnPropertyChanged(nameof(FilterName)); }
            }

            protected void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion

        private void UpdateViewTemplate()
        {
            if (SelViewType != null)
            {
                ViewTemplates = new ObservableCollection<View>(new FilteredElementCollector(ThisDoc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => v.IsTemplate && v.ViewType == SelViewType)
                    .OrderBy(v => v.Name)
                    .ToList());
            }
        }

        private void UpdateFilters()
        {
            if (SelViewTemplate != null)
            {
                Filters = new ObservableCollection<FilterObj>(
                    SelViewTemplate.GetFilters()
                        .Select(id => new FilterObj(ThisDoc, id))
                        .ToList());
            }
        }

        private void UpdateViewTarget()
        {
            ViewTargets = new ObservableCollection<View>(
                new FilteredElementCollector(ThisDoc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate && v.IsViewValidForTemplateCreation() && v.ViewType == SelViewTypeTarget)
                    .OrderBy(v => v.Name)
                    .ToList());
        }

        private void UpdateOptionState()
        {
            PatternCopyEnabled = !AllCopyBOX;
            CutSetCopyEnabled = !AllCopyBOX;
        }
    }
}
