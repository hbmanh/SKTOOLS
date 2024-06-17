using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using System.Collections.ObjectModel;
using System.Windows.Input;
using Application = Autodesk.Revit.ApplicationServices.Application;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Controls;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI.Selection;

namespace SKToolsAddins.ViewModel
{
    public class ChangeBwTypeAndInsViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Application ThisApp;
        private Document ThisDoc;

        public ChangeBwTypeAndInsViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisApp = UiApp.Application;
            ThisDoc = UiDoc.Document;

            InsProParas = new ObservableCollection<ParamObj>();
            SelInsProParas = new ObservableCollection<ParamObj>();
            SelInsProParasToChange = new List<ParamObj>();

            TypProParas = new ObservableCollection<ParamObj>();
            SelTypProParas = new ObservableCollection<ParamObj>();
            SelTypProParasToChange = new List<ParamObj>();

            InsFamParas = new ObservableCollection<ParamObj>();
            SelInsFamParas = new ObservableCollection<ParamObj>();
            SelInsFamParasToChange = new List<ParamObj>();

            TypFamParas = new ObservableCollection<ParamObj>();
            SelTypFamParas = new ObservableCollection<ParamObj>();
            SelTypFamParasToChange = new List<ParamObj>();

            SelFamFamilies = new ObservableCollection<Family>();
            SelProFamilies = new ObservableCollection<Family>();


            Categories = new ObservableCollection<Category>(new FilteredElementCollector(ThisDoc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Select(family => family.FamilyCategory)
                .Where(c => c != null && c.CategoryType == CategoryType.Model)
                .GroupBy(c => c.Id)
                .Select(g => g.First())
                .OrderBy(c => c.Name)
                .ToList());

            SelectedElementIds = UiDoc.Selection.GetElementIds().ToList();
            if (SelectedElementIds.Count == 0)
            {
                SelFamCategory = Categories.FirstOrDefault();
                SelProCategory = Categories.FirstOrDefault();
                UpdateFamilies();
            }
            else
            {
                while (SelectedElementIds.Count > 1)
                {
                    TaskDialog.Show("Error", "Please select exactly one element");
                    var pickedIds = UiDoc.Selection.PickObjects(ObjectType.Element, "Select one element");
                    UiDoc.Selection.SetElementIds(new List<ElementId>());
                    if (pickedIds.Count == 1)
                    {
                        UiDoc.Selection.SetElementIds(new List<ElementId>());
                        SelectedElementIds.Add(pickedIds[0].ElementId);
                    }
                    else
                    {

                    }
                }
                Element element = ThisDoc.GetElement(SelectedElementIds.FirstOrDefault());
                if (element is FamilyInstance familyInstance)
                {
                    Family insFamily = familyInstance.Symbol.Family;
                    SelFamCategory = insFamily.FamilyCategory;
                    SelProCategory = insFamily.FamilyCategory;
                    UpdateFamilies();
                    SelFamFamilies.Add(insFamily);
                    SelProFamilies.Add(insFamily);
                }
            }
            
            UpdateFamParas();
            UpdateProParas();
            SelFamFamilies.CollectionChanged += (sender, args) => UpdateFamParas();

            //SelInsProParas.CollectionChanged += (sender, args) =>
            //{
            //    if (InsProParas != null)
            //        foreach (var selInsProPara in SelInsProParas)
            //        {
            //            var selInsProFam = selInsProPara.ParamValue.Element as Family;
            //            SelProFamilies.Add(selInsProFam);
            //        }
            //};
            //SelTypProParas.CollectionChanged += (sender, args) =>
            //{
            //    if (TypProParas != null) foreach (var selTypProPara in SelTypProParas)
            //        SelProFamilies.Add(selTypProPara.ParamValue.Element as Family);
            //};

            InsToTypFamICommand = new RelayCommand(_InsToTypeFamICommand);
            TypToInsFamICommand = new RelayCommand(_TypToInsFamICommand);

            InsToTypProICommand = new RelayCommand(_InsToTypProICommand);
            TypToInsProICommand = new RelayCommand(_TypToInsProICommand);
        }

        #region Properties

        public List<ParamObj> SelTypFamParasToChange { get; set; }
        public List<ParamObj> SelInsFamParasToChange { get; set; }
        public List<ParamObj> SelInsProParasToChange { get; set; }
        public List<ParamObj> SelTypProParasToChange { get; set; }
        public List<FamilyInstance> FamFamInstances { get; set; }
        public List<FamilyInstance> ProFamInstances { get; set; }

        private ObservableCollection<Family> _famFamilies;

        public ObservableCollection<Family> FamFamilies
        {
            get => _famFamilies;
            set
            {
                _famFamilies = value;
                OnPropertyChanged(nameof(FamFamilies));
            }
        }
        private ObservableCollection<Family> _proFamilies;

        public ObservableCollection<Family> ProFamilies
        {
            get => _proFamilies;
            set
            {
                _proFamilies = value;
                OnPropertyChanged(nameof(ProFamilies));
            }
        }

        private ObservableCollection<Family> _selFamilies { get; set; }

        public ObservableCollection<Family> SelFamFamilies
        {
            get { return _selFamilies; }
            set
            {
                _selFamilies = value;
                OnPropertyChanged(nameof(SelFamFamilies));
            }
        }
        private ObservableCollection<Family> _selProFamilies { get; set; }

        public ObservableCollection<Family> SelProFamilies
        {
            get { return _selProFamilies; }
            set
            {
                _selProFamilies = value;
                OnPropertyChanged(nameof(SelProFamilies));
            }
        }
        private ObservableCollection<ParamObj> _insProParas { get; set; }

        public ObservableCollection<ParamObj> InsProParas
        {
            get { return _insProParas; }
            set
            {
                _insProParas = value;
                OnPropertyChanged(nameof(InsProParas));
            }
        }

        private ObservableCollection<ParamObj> _selInsProParas { get; set; }

        public ObservableCollection<ParamObj> SelInsProParas
        {
            get { return _selInsProParas; }
            set
            {
                _selInsProParas = value;
                OnPropertyChanged(nameof(SelInsProParas));
                if (InsProParas != null)
                    foreach (var selInsProPara in SelInsProParas)
                    {
                        var selInsProFam = selInsProPara.ParamValue.Element as Family;
                        SelProFamilies.Add(selInsProFam);
                    }
            }
        }

        private ObservableCollection<ParamObj> _typProParas { get; set; }

        public ObservableCollection<ParamObj> TypProParas
        {
            get { return _typProParas; }
            set
            {
                _typProParas = value;
                OnPropertyChanged(nameof(TypProParas));
            }
        }

        private ObservableCollection<ParamObj> _selTypProParas { get; set; }

        public ObservableCollection<ParamObj> SelTypProParas
        {
            get { return _selTypProParas; }
            set
            {
                _selTypProParas = value;
                OnPropertyChanged(nameof(SelTypProParas));
                if (TypProParas != null)
                    foreach (var selTypProPara in SelTypProParas)
                    {
                        var selInsProFam = selTypProPara.ParamValue.Element as Family;
                        SelProFamilies.Add(selInsProFam);
                    }
            }
        }

        private ObservableCollection<ParamObj> _insFamParas { get; set; }

        public ObservableCollection<ParamObj> InsFamParas
        {
            get { return _insFamParas; }
            set
            {
                _insFamParas = value;
                OnPropertyChanged(nameof(InsFamParas));
            }
        }

        private ObservableCollection<ParamObj> _selInsFamParas { get; set; }

        public ObservableCollection<ParamObj> SelInsFamParas
        {
            get { return _selInsFamParas; }
            set
            {
                _selInsFamParas = value;
                OnPropertyChanged(nameof(SelInsFamParas));
            }
        }

        private ObservableCollection<ParamObj> _insFamParasToChange { get; set; }

        public ObservableCollection<ParamObj> InsFamParasToChange
        {
            get { return _insFamParasToChange; }
            set
            {
                _insFamParasToChange = value;
                OnPropertyChanged(nameof(InsFamParasToChange));
            }
        }

        private ObservableCollection<ParamObj> _typFamParas { get; set; }

        public ObservableCollection<ParamObj> TypFamParas
        {
            get { return _typFamParas; }
            set
            {
                _typFamParas = value;
                OnPropertyChanged(nameof(TypFamParas));
            }
        }

        private ObservableCollection<ParamObj> _selTypFamParas { get; set; }

        public ObservableCollection<ParamObj> SelTypFamParas
        {
            get { return _selTypFamParas; }
            set
            {
                _selTypFamParas = value;
                OnPropertyChanged(nameof(SelTypFamParas));
            }
        }

        private ObservableCollection<Category> _famCategories { get; set; }

        public ObservableCollection<Category> Categories
        {
            get { return _famCategories; }
            set
            {
                _famCategories = value;
                OnPropertyChanged(nameof(Categories));
            }
        }

        private List<ElementId> _selectedElementIds { get; set; }

        public List<ElementId> SelectedElementIds
        {
            get { return _selectedElementIds; }
            set
            {
                _selectedElementIds = value;
                OnPropertyChanged(nameof(SelectedElementIds));
            }
        }

        private Category _SelFamCategory { get; set; }

        public Category SelFamCategory
        {
            get { return _SelFamCategory; }
            set
            {
                _SelFamCategory = value;
                OnPropertyChanged(nameof(SelFamCategory));
                UpdateFamilies();
                InsFamParas.Clear();
                TypFamParas.Clear();
            }
        }
        public Category _selProCategory { get; set; }

        public Category SelProCategory
        {
            get { return _selProCategory; }
            set
            {
                _selProCategory = value;
                OnPropertyChanged(nameof(SelProCategory));
                UpdateFamilies();
                InsProParas.Clear();
                TypProParas.Clear();
                UpdateProParas();
            }
        }
        public ICommand InsToTypFamICommand { get; set; }

        public void _InsToTypeFamICommand(object Obj)
        {
            foreach (var selInsFamPara in SelInsFamParas.ToList())
            {
                TypFamParas.Add(selInsFamPara);
                InsFamParas.Remove(selInsFamPara);
                SelInsFamParasToChange.Add(selInsFamPara);
            }
        }

        public ICommand TypToInsFamICommand { get; set; }

        public void _TypToInsFamICommand(object Obj)
        {
            foreach (var selTypeFamPara in SelTypFamParas.ToList())
            {
                InsFamParas.Add(selTypeFamPara);
                TypFamParas.Remove(selTypeFamPara);
                SelTypFamParasToChange.Add(selTypeFamPara);
            }
        }
        public ICommand InsToTypProICommand { get; set; }

        public void _InsToTypProICommand(object Obj)
        {
            foreach (var selInsProPara in SelInsProParas.ToList())
            {
                TypProParas.Add(selInsProPara);
                InsProParas.Remove(selInsProPara);
                SelInsProParasToChange.Add(selInsProPara);
            }
        }
        public ICommand TypToInsProICommand { get; set; }

        public void _TypToInsProICommand(object Obj)
        {
            foreach (var selTypProPara in SelTypProParas.ToList())
            {
                InsProParas.Add(selTypProPara);
                TypProParas.Remove(selTypProPara);
                SelTypProParasToChange.Add(selTypProPara);
            }
        }

        public class ParamObj : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            public ParamObj(Parameter parameter)
            {
                Param = parameter;
                ParamName = parameter.Definition.Name;
                ParamType = parameter.Definition.ParameterType;
                ParamGroup = parameter.Definition.ParameterGroup;
                ParamValue = parameter;

            }

            public ParamObj(FamilyParameter familyParameter)
            {
                FamPara = familyParameter;
                ParamName = familyParameter.Definition.Name;
                ParamType = familyParameter.Definition.ParameterType;
                ParamGroup = familyParameter.Definition.ParameterGroup;
                ParamValue = familyParameter;
                ParaIsInstance = familyParameter.IsInstance;
            }

            private Parameter param;

            public Parameter Param
            {
                get { return param; }
                set
                {
                    param = value;
                    OnPropertyChanged(nameof(Param));
                    ParamName = param.Definition.Name;
                    ParamValue = param;
                }
            }

            private FamilyParameter famPara { get; set; }

            public FamilyParameter FamPara
            {
                get { return famPara; }
                set
                {
                    famPara = value;
                    OnPropertyChanged(nameof(FamPara));
                    ParamName = famPara.Definition.Name;
                    ParamType = famPara.Definition.ParameterType;
                    ParamGroup = famPara.Definition.ParameterGroup;
                    ParamValue = famPara;
                }
            }

            private string paramName;

            public string ParamName
            {
                get { return paramName; }
                set
                {
                    paramName = value;
                    OnPropertyChanged(nameof(ParamName));
                }
            }

            private ParameterType paramType;

            public ParameterType ParamType
            {
                get { return paramType; }
                set
                {
                    paramType = value;
                    OnPropertyChanged(nameof(ParamType));
                }
            }

            private BuiltInParameterGroup paramGroup;

            public BuiltInParameterGroup ParamGroup
            {
                get { return paramGroup; }
                set
                {
                    paramGroup = value;
                    OnPropertyChanged(nameof(ParamGroup));
                }
            }

            private dynamic paramValue;

            public dynamic ParamValue
            {
                get { return paramValue; }
                set
                {
                    paramValue = value;
                    OnPropertyChanged(nameof(ParamValue));
                }
            }

            private bool paraIsInstance;

            public bool ParaIsInstance
            {
                get { return paraIsInstance; }
                set
                {
                    paraIsInstance = value;
                    OnPropertyChanged(nameof(ParaIsInstance));
                }
            }

            protected void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion
        
        private void UpdateFamilies()
        {
            if (SelFamCategory != null)
            {
                FamFamilies = new ObservableCollection<Family>(new FilteredElementCollector(ThisDoc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .Where(family => family.FamilyCategory.Id == SelFamCategory.Id)
                    .OrderBy(family => family.Name)
                    .ToList());
            }

            if (SelProCategory != null)
            {
                ProFamilies = new ObservableCollection<Family>(new FilteredElementCollector(ThisDoc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .Where(family => family.FamilyCategory.Id == SelProCategory.Id)
                    .OrderBy(family => family.Name)
                    .ToList());
            }
        }

        private void UpdateFamParas()
        {
            InsFamParas.Clear();
            TypFamParas.Clear();

            List<ParamObj> parameters = new List<ParamObj>();
            if (SelFamFamilies.Count != 0)
            {
                foreach (var selFamily in SelFamFamilies)
                {
                    var famDoc = ThisDoc.EditFamily(selFamily);
                    var famManager = famDoc.FamilyManager;
                    var famParas = famManager.Parameters;

                    foreach (FamilyParameter famPara in famParas)
                    {
                        ParamObj famParaObj = new ParamObj(famPara);
                        Debug.Print(famParaObj.ParamName);
                        parameters.Add(famParaObj);
                    }

                    FamFamInstances = new FilteredElementCollector(ThisDoc)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(i => i.Symbol.Family.Id == selFamily.Id)
                        .ToList();

                    if (FamFamInstances != null)
                    {
                        foreach (var instance in FamFamInstances)
                        {
                            var iList = instance.GetOrderedParameters();
                            foreach (var iLis in iList) parameters.Add(new ParamObj(iLis));
                        }
                    }
                }

                var paraObjeOrderBy = parameters.GroupBy(p => p.ParamName)
                    .Select(g => g.FirstOrDefault()).OrderBy(n => n.ParamName);
                foreach (var p in paraObjeOrderBy)
                {
                    if (p.ParaIsInstance && !p.ParamValue.IsReadOnly)
                        InsFamParas.Add(p);
                    if (!p.ParaIsInstance && !p.ParamValue.IsReadOnly)
                        TypFamParas.Add(p);
                }
            }
        }
        private void UpdateProParas()
        {
            InsProParas.Clear();
            TypProParas.Clear();
           
            List<ParamObj> insProParamObj = new List<ParamObj>();
            List<ParamObj> typProParamObj = new List<ParamObj>();

            ProFamInstances = new FilteredElementCollector(ThisDoc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(i => i.Symbol.Category.Id == SelProCategory.Id)
                .ToList();

            if (ProFamInstances != null)
            {
                TypProParas.Clear();
                InsProParas.Clear();

                foreach (var instance in ProFamInstances)
                {
                    var insParams = instance.GetOrderedParameters();
                    var typParams = instance.Symbol.GetOrderedParameters();

                    foreach (var insParam in insParams)
                        if (!insParam.IsReadOnly) insProParamObj.Add(new ParamObj(insParam));

                    foreach (var typParam in typParams)
                        if (!typParam.IsReadOnly) typProParamObj.Add(new ParamObj(typParam));

                    var famDoc = ThisDoc.EditFamily(instance.Symbol.Family);
                    var famManager = famDoc.FamilyManager;
                    var famParas = famManager.Parameters;
                    foreach (FamilyParameter famPara in famParas)
                    {
                        if (!famPara.IsReadOnly)
                        {
                            ParamObj famParaObj = new ParamObj(famPara);
                            insProParamObj.RemoveAll(paramObj => paramObj.ParamName == famParaObj.ParamName);
                            typProParamObj.RemoveAll(paramObj => paramObj.ParamName == famParaObj.ParamName);
                        }
                    }
                }
                var insParamObjOrderBy = insProParamObj.GroupBy(p => p.ParamName)
                    .Select(g => g.FirstOrDefault()).OrderBy(n => n.ParamName);

                if (insParamObjOrderBy != null)
                    foreach (var insParamObjOrder in insParamObjOrderBy)
                        InsProParas.Add(insParamObjOrder);

                var typParamObjOrderBy = typProParamObj.GroupBy(p => p.ParamName)
                    .Select(g => g.FirstOrDefault()).OrderBy(n => n.ParamName);

                if (typParamObjOrderBy != null)
                    foreach (var typParamObjOrder in typParamObjOrderBy)
                        TypProParas.Add(typParamObjOrder);
            }
        }
    }
}
