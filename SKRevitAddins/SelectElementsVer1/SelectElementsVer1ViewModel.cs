using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKRevitAddins.SelectElementsVer1;
using SKRevitAddins.Utils;
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace SKRevitAddins.ViewModel
{
    public class SelectElementsVer1ViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Application ThisApp;
        private Document ThisDoc;

        public SelectElementsVer1ViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisApp = UiApp.Application;
            ThisDoc = UiDoc.Document;

            FilterBy = new List<string>()
            {
                //"Active View",
                //"Level",
                //"Current Selection",
                //"Entire Project",
                "Numbering",
                "Replace symbol",
            };
            SelectedElementIds = UiDoc.Selection.GetElementIds().ToList();
            if (SelectedElementIds.Count > 0)
            {
                SelFilterBy = "Replace symbol";
            }
            else
            {
                TaskDialog.Show("Error", "Please select elements before running add-in");
            }
            //SelFilterBy = SelectedElementIds.Count > 0 ? "Current Selection" : "Active View";
            GetAllElementsBasedOnSelectedElements();

            //UpdateLevels();

            Categories = new ObservableCollection<Category>();
            GetCategoryBasedOnSelectedElements();

            Parameters = new ObservableCollection<Parameter>();
            UpdateParaFromSelElements();

            StringRules = new List<string>()
            {
                "equals",
                "does not equal",
                //"is greater than",
                //"is greater than or equal to",
                //"is less than",
                //"is less than or equal to",
                "contains",
                "does not contain",
                "begins with",
                "does not begin with",
                "ends with",
                "does not end with"
            };
            SelStringRule = StringRules.FirstOrDefault();

            GetParamValuesFromSelPara();
            ParamValue = ParamValues.FirstOrDefault();

            UpdateSelElements();
            //SelEleToPreview = EleToPreview.FirstOrDefault();
            if (SelFilterBy == "Numbering") ValSetICommand = new RelayCommand(_ValSetICommand);
            if (SelFilterBy == "Replace symbol") ReplaceTextICommand = new RelayCommand(_ReplaceTextICommand);

            IsLeftToRight = true;
            IsUpToDown = true;
        }

        #region Properties
        private List<FamilyInstance> _elementsSelected;

        public List<FamilyInstance> ElementsSelected
        {
            get { return _elementsSelected; }
            set
            {
                _elementsSelected = value;
                OnPropertyChanged(nameof(ElementsSelected));
            }
        }

        private List<string> _filterBy;

        public List<string> FilterBy
        {
            get { return _filterBy; }
            set
            {
                _filterBy = value;
                OnPropertyChanged(nameof(FilterBy));
            }
        }

        private string _selFilterBy;

        public string SelFilterBy
        {
            get { return _selFilterBy; }
            set
            {
                _selFilterBy = value;
                OnPropertyChanged(nameof(SelFilterBy));
                //UpdateLevels();
                GetCategoryBasedOnSelectedElements();
            }
        }
        private List<ElementId> selectedElementIds { get; set; }
        public List<ElementId> SelectedElementIds
        {
            get { return selectedElementIds; }
            set { selectedElementIds = value; OnPropertyChanged(nameof(SelectedElementIds)); }
        }

        private List<Level> _levels;

        public List<Level> Levels
        {
            get { return _levels; }
            set
            {
                _levels = value;
                OnPropertyChanged(nameof(Levels));
            }
        }

        private Level _selLevel;

        public Level SelLevel
        {
            get { return _selLevel; }
            set
            {
                _selLevel = value;
                OnPropertyChanged(nameof(SelLevel));
                //GetCategoryBasedOnSelectedElements();
            }
        }

        private ObservableCollection<Category> _categories;

        public ObservableCollection<Category> Categories
        {
            get { return _categories; }
            set { _categories = value; OnPropertyChanged(nameof(Categories)); }
        }
        private Category _selCategory;

        public Category SelCategory
        {
            get { return _selCategory; }
            set
            {
                _selCategory = value;
                OnPropertyChanged(nameof(SelCategory));
                UpdateParaFromSelElements();
                UpdateSelElements();
            }
        }

        private ObservableCollection<Parameter> _parameters;

        public ObservableCollection<Parameter> Parameters
        {
            get { return _parameters; }
            set
            { _parameters = value; OnPropertyChanged(nameof(Parameters)); }
        }

        private Parameter _selParameter;

        public Parameter SelParameter
        {
            get { return _selParameter; }
            set
            {
                _selParameter = value;
                OnPropertyChanged(nameof(SelParameter));
                GetParamValuesFromSelPara();
                UpdateSelElements();
            }
        }

        private bool _isLevelEnable;

        public bool IsLevelEnable
        {
            get { return _isLevelEnable; }
            set { _isLevelEnable = value; OnPropertyChanged(nameof(IsLevelEnable)); }
        }

        private List<string> _stringRules;

        public List<string> StringRules
        {
            get { return _stringRules; }
            set { _stringRules = value; OnPropertyChanged(nameof(StringRules)); }
        }
        private string _selStringRule;

        public string SelStringRule
        {
            get { return _selStringRule; }
            set
            {
                _selStringRule = value;
                OnPropertyChanged(nameof(SelStringRule));
                UpdateSelElements();
            }
        }
        public ICommand ValSetICommand { get; set; }
        public void _ValSetICommand(object obj)
        {
            SelectElementsVer1NumberingRuleWpfWindow selectElementsVer1NumberingRuleWpfWindow = new SelectElementsVer1NumberingRuleWpfWindow(this);
            selectElementsVer1NumberingRuleWpfWindow.ShowDialog();
        }
        public ICommand ReplaceTextICommand { get; set; }
        public void _ReplaceTextICommand(object obj)
        {
            SelectElementsVer1ReplaceTextWpfWindow selectElementsVer1ReplaceTextWpfWindow = new SelectElementsVer1ReplaceTextWpfWindow(this);
            selectElementsVer1ReplaceTextWpfWindow.ShowDialog();
        }
        private ObservableCollection<object> _paramValues;

        public ObservableCollection<object> ParamValues
        {
            get { return _paramValues; }
            set { _paramValues = value; OnPropertyChanged(nameof(ParamValues)); }
        }

        private object _paramValue;

        public object ParamValue
        {
            get { return _paramValue; }
            set
            {
                _paramValue = value;
                OnPropertyChanged(nameof(ParamValue));
            }
        }
        private List<Element> _eleToPreview;

        public List<Element> EleToPreview
        {
            get { return _eleToPreview; }
            set { _eleToPreview = value; OnPropertyChanged(nameof(EleToPreview)); }
        }

        //private Element _selEleToPreview;

        //public Element SelEleToPreview
        //{
        //    get { return _selEleToPreview; }
        //    set { _selEleToPreview = value; OnPropertyChanged(nameof(SelEleToPreview)); }
        //}
        private bool _isUpToDown { get; set; }
        public bool IsUpToDown
        {
            get { return _isUpToDown; }
            set { _isUpToDown = value; OnPropertyChanged(nameof(IsUpToDown)); }
        }
        private bool _isDownToUp { get; set; }
        public bool IsDownToUp
        {
            get { return _isDownToUp; }
            set { _isDownToUp = value; OnPropertyChanged(nameof(IsDownToUp)); }
        }

        private bool isLeftToRight { get; set; }
        public bool IsLeftToRight
        {
            get { return isLeftToRight; }
            set { isLeftToRight = value; OnPropertyChanged(nameof(IsLeftToRight)); }
        }

        private bool isRightToLeft { get; set; }

        public bool IsRightToLeft
        {
            get { return isRightToLeft; }
            set { isRightToLeft = value; OnPropertyChanged(nameof(IsRightToLeft)); }
        }

        private string _beginsWith { get; set; }
        public string BeginsWith
        {
            get { return _beginsWith; }
            set { _beginsWith = value; OnPropertyChanged(nameof(BeginsWith)); }
        }

        private string _keywords { get; set; }
        public string Keywords
        {
            get { return _keywords; }
            set { _keywords = value; OnPropertyChanged(nameof(Keywords)); }
        }
        private string _keyTarget { get; set; }
        public string KeyTarget
        {
            get { return _keyTarget; }
            set { _keyTarget = value; OnPropertyChanged(nameof(KeyTarget)); }
        }

        #endregion


        //private void UpdateLevels()
        //{
        //    if (SelFilterBy != "Level") IsLevelEnable = false; else IsLevelEnable = true;

        //    Levels = new FilteredElementCollector(ThisDoc)
        //        .OfCategory(BuiltInCategory.OST_Levels)
        //        .OfClass(typeof(Level))
        //        .Cast<Level>()
        //        .OrderBy(n => n.Name)
        //        .ToList();
        //    SelLevel = Levels.FirstOrDefault();
        //}

        private void GetAllElementsBasedOnSelectedElements()
        {
            //if (SelFilterBy == "Active View")
            //{
            //    ElementsSelected =new FilteredElementCollector(ThisDoc, ThisDoc.ActiveView.Id)
            //        .OfClass(typeof(FamilyInstance))
            //        .Cast<FamilyInstance>()
            //        .ToList();
            //}

            //if (SelFilterBy == "Level" && SelLevel != null)
            //{
            //    ElementsSelected = new FilteredElementCollector(ThisDoc)
            //        .OfCategoryId(SelLevel.Id)
            //        .WhereElementIsNotElementType()
            //        .ToList();
            //}
            ElementsSelected = new FilteredElementCollector(ThisDoc, UiDoc.Selection.GetElementIds())
                .WhereElementIsNotElementType()
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();
            //if (SelFilterBy == "Current Selection")
            //{
            //    ElementsSelected = new FilteredElementCollector(ThisDoc, UiDoc.Selection.GetElementIds())
            //        .WhereElementIsNotElementType()
            //        .OfClass(typeof(FamilyInstance))
            //        .Cast<FamilyInstance>()
            //        .ToList();
            //}

            //if (SelFilterBy == "Entire Project")
            //{

            //}
        }
        private void GetCategoryBasedOnSelectedElements()
        {
            Categories = new ObservableCollection<Category>();
            Categories.Clear();
            try
            {
                Categories = new ObservableCollection<Category>(ElementsSelected.Select(element => element.Category)
                    .Where(c => c != null)
                    .GroupBy(c => c.Id)
                    .Select(c => c.First())
                    .OrderBy(c => c.Name)
                    .ToList());

                SelCategory = Categories.FirstOrDefault();
            }
            catch (Exception e)
            {
                //TaskDialog.Show("エラー", "要素がありません");
            }
        }

        private void UpdateParaFromSelElements()
        {
            Parameters = new ObservableCollection<Parameter>();
            Parameters.Clear();
            Parameters = new ObservableCollection<Parameter>(new FilteredElementCollector(ThisDoc)
                .OfCategoryId(SelCategory.Id)
                .SelectMany(element => element.Parameters.Cast<Parameter>())
                .GroupBy(p => p.Definition.Name)
                .Select(p => p.First())
                .OrderBy(p => p.Definition.Name)
                .ToList());
            SelParameter = Parameters.FirstOrDefault();
        }

        private void GetParamValuesFromSelPara()
        {
            ParamValues = new ObservableCollection<object>();
            ParamValues.Clear();

            foreach (var element in ElementsSelected)
            {
                if (SelParameter == null) continue;
                var parameter = element.LookupParameter(SelParameter.Definition.Name);

                if (parameter == null) continue;
                var paramValueString = parameter.AsValueString();
                if (!string.IsNullOrEmpty(paramValueString) && ParamValues.All(p => p.ToString() != paramValueString)) ParamValues.Add(paramValueString);
            }
            ParamValues = new ObservableCollection<object>(ParamValues.Distinct().OrderBy(n => n));
            ParamValue = ParamValues.FirstOrDefault();
        }

        private void UpdateSelElements()
        {
            EleToPreview = new List<Element>();
            EleToPreview.Clear();
            foreach (var element in ElementsSelected)
            {
                if (SelParameter == null || SelCategory.Id != element.Category.Id) continue;
                var parameter = element.LookupParameter(SelParameter.Definition.Name);
                if (parameter == null || ParamValue == null) continue;
                var paramValueString = ParamValue.ToString();
                if (EleToPreview.All(e => e.Id != element.Id) && IsMatch(paramValueString, SelStringRule))
                {
                    EleToPreview.Add(element);
                }

            }
            var eleCount = EleToPreview.Count;
        }


        private bool IsMatch(string parameterValue, string rule)
        {
            switch (rule)
            {
                case "equals":
                    return parameterValue == ParamValue;
                case "does not equal":
                    return parameterValue != ParamValue;
                //case "is greater than":
                //    if (double.TryParse(parameterValue, out double value1) 
                //        && double.TryParse(ParamValue.ToString(), out double value2))
                //    {
                //        return value1 > value2;
                //    }
                //    return false;
                //case "is greater than or equal":
                //    if (double.TryParse(parameterValue, out double value3) 
                //        && double.TryParse(ParamValue.ToString(), out double value4))
                //    {
                //        return value3 >= value4;
                //    }
                //    return false;
                //case "is less than":
                //    if (double.TryParse(parameterValue, out double value5) 
                //        && double.TryParse(ParamValue.ToString(), out double value6))
                //    {
                //        return value5 < value6;
                //    }
                //    return false;
                //case "is less than or equal":
                //    if (double.TryParse(parameterValue, out double value7) 
                //        && double.TryParse(ParamValue.ToString(), out double value8))
                //    {
                //        return value7 <= value8;
                //    }
                //    return false;
                case "contains":
                    return parameterValue.Contains(ParamValue.ToString());
                case "does not contain":
                    return !parameterValue.Contains(ParamValue.ToString());
                case "begins with":
                    return parameterValue.StartsWith(ParamValue.ToString());
                case "does not begin with":
                    return !parameterValue.StartsWith(ParamValue.ToString());
                case "ends with":
                    return parameterValue.EndsWith(ParamValue.ToString());
                case "does not end with":
                    return !parameterValue.EndsWith(ParamValue.ToString());
                default:
                    return false;
            }
        }


    }
}
