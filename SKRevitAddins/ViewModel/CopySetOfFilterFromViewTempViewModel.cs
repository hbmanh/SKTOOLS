using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace SKRevitAddins.ViewModel
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
                    .Select(v => v.ViewType )
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

        private ObservableCollection<ViewType> _viewTypes { get; set; }
        public ObservableCollection<ViewType> ViewTypes
        {
            get { return _viewTypes; }
            set
            {
                _viewTypes = value;
                OnPropertyChanged(nameof(ViewTypes));
            }
        }
        private ViewType _selViewType;
        public ViewType SelViewType
        {
            get { return _selViewType; }
            set
            {
                _selViewType = value;
                OnPropertyChanged(nameof(SelViewType));
                UpdateViewTemplate();
            }
        }
        private ObservableCollection<ViewType> _viewsTypeTarget { get; set; }
        public ObservableCollection<ViewType> ViewsTypeTarget
        {
            get { return _viewsTypeTarget; }
            set
            {
                _viewsTypeTarget = value;
                OnPropertyChanged(nameof(ViewsTypeTarget));
            }
        }
        private ObservableCollection<View> _viewTemplates { get; set; }

        public ObservableCollection<View> ViewTemplates
        {
            get { return _viewTemplates; }
            set
            {
                _viewTemplates = value;
                OnPropertyChanged(nameof(ViewTemplates));
            }
        }

        private View _selViewTemplate;
        public View SelViewTemplate
        {
            get { return _selViewTemplate; }
            set
            {
                _selViewTemplate = value;
                OnPropertyChanged(nameof(SelViewTemplate));
                UpdateFilters();
                //UpdateViewTarget();
            }
        }
        private ObservableCollection<FilterObj> _filters { get; set; }

        public ObservableCollection<FilterObj> Filters
        {
            get { return _filters; }
            set
            {
                _filters = value;
                OnPropertyChanged(nameof(Filters));
            }
        }
        private ObservableCollection<FilterObj> _selFilter;

        public ObservableCollection<FilterObj> SelFilter
        {
            get { return _selFilter; }
            set
            {
                _selFilter = value;
                OnPropertyChanged(nameof(SelFilter));
            }
        }
        private ObservableCollection<View> _viewTargets { get; set; }
        public ObservableCollection<View> ViewTargets
        {
            get { return _viewTargets; }
            set
            {
                _viewTargets = value;
                OnPropertyChanged(nameof(ViewTargets));
            }
        }

        private ObservableCollection<View> _selViewTarget;
        public ObservableCollection<View> SelViewTarget
        {
            get { return _selViewTarget; }
            set
            {
                _selViewTarget = value;
                OnPropertyChanged(nameof(SelViewTarget));
            }
        }

        private ViewType _selViewTypeTarget;

        public ViewType SelViewTypeTarget
        {
            get { return _selViewTypeTarget; }
            set
            {
                _selViewTypeTarget = value;
                OnPropertyChanged(nameof(SelViewTypeTarget));
                UpdateViewTarget();
            }
        }
        public bool _allCopyBOX;

        public bool AllCopyBOX
        {
            get { return _allCopyBOX; }
            set
            {
                _allCopyBOX = value;
                OnPropertyChanged(nameof(AllCopyBOX));
            }
        }

        public bool _patternCopyBOX;

        public bool PatternCopyBOX
        {
            get { return _patternCopyBOX; }
            set
            {
                _patternCopyBOX = value;
                OnPropertyChanged(nameof(PatternCopyBOX));
            }
        }

        public bool _cutSetCopyBOX;

        public bool CutSetCopyBOX
        {
            get { return _cutSetCopyBOX; }
            set
            {
                _cutSetCopyBOX = value;
                OnPropertyChanged(nameof(CutSetCopyBOX));
            }
        }

        public class FilterObj : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            public FilterObj(Document doc,ElementId filterElementId)
            {
                FilterId = filterElementId;
                FilterName = doc.GetElement(filterElementId).Name;
            }
            public bool IsMatch(View view)
            {
                return view.Name.Contains(this.FilterName);
            }
            private ElementId _filterId;

            public ElementId FilterId
            {
                get { return _filterId; }
                set
                {
                    _filterId = value;
                    OnPropertyChanged(nameof(FilterId));
                
                }
            }

            private string _filterName;

            public string FilterName
            {
                get { return _filterName; }
                set
                {
                    _filterName = value;
                    OnPropertyChanged(nameof(FilterName));
                }
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

                //var templatesForSelectedType = new FilteredElementCollector(ThisDoc)
                //    .OfClass(typeof(View))
                //    .Cast<View>()
                //    .Where(v => v.IsTemplate /*&& v.ViewType == SelViewType.ViewFamily*/)
                //    .ToList();


                //if (templatesForSelectedType.Count > 0)
                //{
                //    Filters = new ObservableCollection<FilterObj>(
                //        templatesForSelectedType[0].GetFilters()
                //            .Select(id => new FilterObj(ThisDoc, id)).ToList());
                //}
                //else
                //{
                //    Filters.Clear();
                //}
            }
        }
        private void UpdateFilters()
        {
            if (SelViewTemplate != null)
            {
                Filters = new ObservableCollection<FilterObj>(SelViewTemplate.GetFilters().Select(id => new FilterObj(ThisDoc,id)).ToList());
            }
        }

        private void UpdateViewTarget()
        {
            ViewTargets = new ObservableCollection<View>(new FilteredElementCollector(ThisDoc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && v.IsViewValidForTemplateCreation() && v.ViewType == SelViewTypeTarget)
                .OrderBy(n => n.Name)
                .ToList());

            var viewColl = new FilteredElementCollector(ThisDoc)
                .OfClass(typeof(View))
                .Cast<View>()
                .ToList();
            foreach (var view in viewColl)
            {
                Debug.WriteLine(view.Category);
            }

            ///Test

            //var viewColl = new FilteredElementCollector(ThisDoc)
            //    .OfClass(typeof(View))
            //    .Cast<View>()
            //    .Where(v => !v.IsTemplate && v.IsViewValidForTemplateCreation() && v.ViewType == SelViewTypeTarget)
            //    .OrderBy(n => n.Name)
            //    .ToList();
            //foreach (var view in viewColl)
            //{
            //    string[] viewGroup = view.Title.Split(':');
            //}
        }

        /// Get target view based on Selected View Template
        //private void UpdateViewTarget()
        //{
        //    if (SelViewTemplate != null)
        //    {
        //        var targetType = SelViewTemplate.ViewType;

        //        ViewTargets = new ObservableCollection<View>(new FilteredElementCollector(ThisDoc)
        //            .OfClass(typeof(View))
        //            .Cast<View>()
        //            .Where(v => !v.IsTemplate && v.ViewType == targetType && v.IsViewValidForTemplateCreation())
        //            .OrderBy(n => n.Name)
        //            .ToList());
        //    }
        //}
    }
}
