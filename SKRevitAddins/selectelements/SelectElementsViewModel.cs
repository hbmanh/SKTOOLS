using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace SKRevitAddins.ViewModel
{
    public class SelectElementsViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Application ThisApp;
        private Document ThisDoc;

        public SelectElementsViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisApp = UiApp.Application;
            ThisDoc = UiDoc.Document;

            FilterBy = new List<string>();
            AddFilterByList();
            SelFilterBy = FilterBy.FirstOrDefault();
            GetAllElementsBasedOnFilterBy();

            UpdateLevels();

            Categories = new ObservableCollection<Category>();
            UpdateCategoryBasedOnFilterBy();

            Parameters = new ObservableCollection<Parameter>();
            UpdateParaFromCategories();

            StringRules = new List<string>();
            AddListStringRule();
            SelStringRule = StringRules.FirstOrDefault();

            GetParamValuesFromSelPara();
            SelParamValue = ParamValues.FirstOrDefault();

            UpdateSelElements();
            SelEleToPreview = EleToPreview.FirstOrDefault();


        }

        #region Properties
        private List<Element> _allEleBasedOnFilterBy;

        public List<Element> AllEleBasedOnFilterBy
        {
            get { return _allEleBasedOnFilterBy; }
            set { _allEleBasedOnFilterBy = value; OnPropertyChanged(nameof(AllEleBasedOnFilterBy)); }
        }

        private List<string> _filterBy;

        public List<string> FilterBy
        {
            get { return _filterBy; }
            set { _filterBy = value; OnPropertyChanged(nameof(FilterBy)); }
        }

        private string _selFilterBy;

        public string SelFilterBy
        {
            get { return _selFilterBy; }
            set
            {
                _selFilterBy = value;
                OnPropertyChanged(nameof(SelFilterBy));
                UpdateLevels();
                UpdateCategoryBasedOnFilterBy();
            }
        }

        private List<Level> _levels;

        public List<Level> Levels
        {
            get { return _levels; }
            set
            { _levels = value; OnPropertyChanged(nameof(Levels)); }
        }

        private Level _selLevel;

        public Level SelLevel
        {
            get { return _selLevel; }
            set
            {
                _selLevel = value;
                OnPropertyChanged(nameof(SelLevel));
                UpdateCategoryBasedOnFilterBy();
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
                UpdateParaFromCategories();
            }
        }

        private ObservableCollection<Parameter> _parameters;

        public ObservableCollection<Parameter> Parameters
        {
            get { return _parameters; }
            set { _parameters = value; OnPropertyChanged(nameof(Parameters)); }
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
            set { _selStringRule = value; OnPropertyChanged(nameof(SelStringRule)); }
        }

        private ObservableCollection<object> _paramValues;

        public ObservableCollection<object> ParamValues
        {
            get { return _paramValues; }
            set
            { _paramValues = value; OnPropertyChanged(nameof(ParamValues)); }
        }

        private object _selParamValue;

        public object SelParamValue
        {
            get { return _selParamValue; }
            set
            {
                _selParamValue = value;
                OnPropertyChanged(nameof(SelParamValue));
                UpdateSelElements();
            }
        }
        private ObservableCollection<Element> _eleToPreview;

        public ObservableCollection<Element> EleToPreview
        {
            get { return _eleToPreview; }
            set { _eleToPreview = value; OnPropertyChanged(nameof(EleToPreview)); }
        }

        private Element _selEleToPreview;

        public Element SelEleToPreview
        {
            get { return _selEleToPreview; }
            set { _selEleToPreview = value; OnPropertyChanged(nameof(SelEleToPreview)); }
        }


        #endregion

        private void AddFilterByList()
        {
            FilterBy.Add("現在ビュー");
            FilterBy.Add("レベル");
            FilterBy.Add("現在の選択");
            FilterBy.Add("プロジェクト全て");
        }

        private void UpdateLevels()
        {
            if (SelFilterBy != "レベル") IsLevelEnable = false; else IsLevelEnable = true;

            Levels = new FilteredElementCollector(ThisDoc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(n => n.Name)
                .ToList();
            SelLevel = Levels.FirstOrDefault();
        }

        private void GetAllElementsBasedOnFilterBy()
        {
            AllEleBasedOnFilterBy = new List<Element>();
            AllEleBasedOnFilterBy.Clear();
            if (SelFilterBy == "現在ビュー")
            {
                AllEleBasedOnFilterBy = new FilteredElementCollector(ThisDoc, ThisDoc.ActiveView.Id)
                    .WhereElementIsNotElementType()
                    .ToList(); ;
            }

            if (SelFilterBy == "レベル" && SelLevel != null)
            {
                AllEleBasedOnFilterBy = new FilteredElementCollector(ThisDoc)
                    .OfCategoryId(SelLevel.Id)
                    .WhereElementIsNotElementType()
                    .ToList();
            }

            if (SelFilterBy == "現在の選択")
            {
                AllEleBasedOnFilterBy = new FilteredElementCollector(ThisDoc, UiDoc.Selection.GetElementIds())
                    .WhereElementIsNotElementType()
                    .ToList();
            }

            if (SelFilterBy == "プロジェクト全て")
            {
                AllEleBasedOnFilterBy = new FilteredElementCollector(ThisDoc)
                    .WhereElementIsNotElementType()
                    .ToList();
            }
        }
        private void UpdateCategoryBasedOnFilterBy()
        {
            Categories = new ObservableCollection<Category>();
            Categories.Clear();
            try
            {
                Categories = new ObservableCollection<Category>(AllEleBasedOnFilterBy.Select(element => element.Category)
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

        private void UpdateParaFromCategories()
        {
            Parameters = new ObservableCollection<Parameter>();
            Parameters.Clear();
            if (SelCategory != null)
            {
                Parameters = new ObservableCollection<Parameter>(new FilteredElementCollector(ThisDoc)
                    .OfCategoryId(SelCategory.Id)
                    .WhereElementIsNotElementType()
                    .SelectMany(element => element.Parameters.Cast<Parameter>())
                    .GroupBy(p => p.Id)
                    .Select(p => p.First())
                    .OrderBy(p => p.Definition.Name)
                    .ToList());
                SelParameter = Parameters.FirstOrDefault();

            }
        }

        private void AddListStringRule()
        {
            StringRules.Add("等しい");
            StringRules.Add("等しくない");
            StringRules.Add("より大きい");
            StringRules.Add("以上");
            StringRules.Add("より小さい");
            StringRules.Add("以下");
            StringRules.Add("含む");
            StringRules.Add("含まない");
            StringRules.Add("で始まる");
            StringRules.Add("で始まらない");
            StringRules.Add("で終わる");
            StringRules.Add("で終わらない");
            SelStringRule = StringRules.FirstOrDefault();

        }
        private void GetParamValuesFromSelPara()
        {
            ParamValues = new ObservableCollection<object>();
            ParamValues.Clear();

            foreach (var element in AllEleBasedOnFilterBy)
            {
                if (SelParameter != null)
                {
                    var parameter = element.LookupParameter(SelParameter.Definition.Name);

                    if (parameter != null)
                    {
                        var paramValueString = parameter.AsValueString();
                        if (!string.IsNullOrEmpty(paramValueString)) ParamValues.Add(paramValueString);
                    }
                }
            }
            ParamValues = new ObservableCollection<object>(ParamValues.Distinct().OrderBy(n => n));
            SelParamValue = ParamValues.FirstOrDefault();
        }

        private void UpdateSelElements()
        {
            EleToPreview = new ObservableCollection<Element>();
            EleToPreview.Clear();

            foreach (var element in AllEleBasedOnFilterBy)
            {
                if (SelParameter != null)
                {
                    var parameter = element.LookupParameter(SelParameter.Definition.Name);

                    if (parameter != null)
                    {
                        var paramValueString = parameter.AsValueString();
                        if (!string.IsNullOrEmpty(paramValueString) && IsMatch(paramValueString, SelStringRule))
                        {
                            EleToPreview.Add(element);
                        }
                    }
                }
            }


        }

        private bool IsMatch(string parameterValue, string rule)
        {
            switch (rule)
            {
                case "等しい":
                    return parameterValue == SelParamValue;
                case "等しくない":
                    return parameterValue != SelParamValue;
                case "より大きい":
                    if (double.TryParse(parameterValue, out double value1)
                        && double.TryParse(SelParamValue.ToString(), out double value2))
                    {
                        return value1 > value2;
                    }
                    return false;
                case "以上":
                    if (double.TryParse(parameterValue, out double value3)
                        && double.TryParse(SelParamValue.ToString(), out double value4))
                    {
                        return value3 >= value4;
                    }
                    return false;
                case "より小さい":
                    if (double.TryParse(parameterValue, out double value5)
                        && double.TryParse(SelParamValue.ToString(), out double value6))
                    {
                        return value5 < value6;
                    }
                    return false;
                case "以下":
                    if (double.TryParse(parameterValue, out double value7)
                        && double.TryParse(SelParamValue.ToString(), out double value8))
                    {
                        return value7 <= value8;
                    }
                    return false;
                case "含む":
                    return parameterValue.Contains(SelParamValue.ToString());
                case "含まない":
                    return !parameterValue.Contains(SelParamValue.ToString());
                case "で始まる":
                    return parameterValue.StartsWith(SelParamValue.ToString());
                case "で始まらない":
                    return !parameterValue.StartsWith(SelParamValue.ToString());
                case "で終わる":
                    return parameterValue.EndsWith(SelParamValue.ToString());
                case "で終わらない":
                    return !parameterValue.EndsWith(SelParamValue.ToString());
                case "値が含まれています":
                    return parameterValue.IndexOf(SelParamValue.ToString(), StringComparison.OrdinalIgnoreCase) >= 0;
                case "値が含まれていません":
                    return parameterValue.IndexOf(SelParamValue.ToString(), StringComparison.OrdinalIgnoreCase) < 0;
                default:
                    return false;
            }
        }


    }
}
