using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using Parameter = Autodesk.Revit.DB.Parameter;
using Application = Autodesk.Revit.ApplicationServices.Application;
using System.Security;
using System.Windows;
using System.Windows.Controls;

namespace ParamCopy
{
    public class ParamCopyViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Application ThisApp;
        private Document ThisDoc;
        private string BoolYes;
        private string BoolNo;

        public ParamCopyViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisApp = UiApp.Application;
            ThisDoc = UiDoc.Document;
            LanguageType appLanguage = ThisApp.Language;
            if (appLanguage == LanguageType.English_USA || appLanguage == LanguageType.English_GB)
            {
                BoolYes = "Yes";
                BoolNo = "No";
            }
            else if (appLanguage == LanguageType.Japanese)
            {
                BoolYes = "はい";
                BoolNo = "いいえ";
            }
            CanCopy = true;
            IsRounded = false;
            RoundValue = 2;

            IsSourceNotFromLink = true;

            SelectSourceICommand = new RelayCommand(_SelectSourceICommand);
            SelectLinkedSourceICommand = new RelayCommand(_SelectLinkedSourceICommand);
            SetTargetSameAsSourceICommand = new RelayCommand(_SetTargetSameAsSourceICommand);
            SelectTargetICommand = new RelayCommand(_SelectTargetICommand);
            RefreshICommand = new RelayCommand(_RefreshICommand);
            InstanceCopyICommand = new RelayCommand(_InstanceCopyICommand);
            FamilyCopyICommand = new RelayCommand(_FamilyCopyICommand);
            CategoryCopyICommand = new RelayCommand(_CategoryCopyICommand);
            AllEleCopyICommand = new RelayCommand(_AllEleCopyICommand);
            SourceInstParams = new ObservableCollection<ParamObj>();
            SourceTypeParams = new ObservableCollection<ParamObj>();
            TargetInstParams = new ObservableCollection<ParamObj>();
            TargetTypeParams = new ObservableCollection<ParamObj>();

            Selection selection = UiDoc.Selection;
            List<ElementId> selectedElementIds = selection.GetElementIds().ToList();
            if (selectedElementIds.Count == 1)
            {
                SourceElement = ThisDoc.GetElement(selectedElementIds.First());
                DefineSourceElementInfo(SourceElement);
                TargetElement = SourceElement;
                UpdateTarget(TargetElement);
            }
        }

        private bool isSourceInstTabEnabled;
        public bool IsSourceInstTabEnabled
        {
            get { return isSourceInstTabEnabled; }
            set
            {
                isSourceInstTabEnabled = value;
                OnPropertyChanged(nameof(IsSourceInstTabEnabled));
            }
        }
        private bool isSourceTypeTabEnabled;
        public bool IsSourceTypeTabEnabled
        {
            get { return isSourceTypeTabEnabled; }
            set
            {
                isSourceTypeTabEnabled = value;
                OnPropertyChanged(nameof(IsSourceTypeTabEnabled));
            }
        }
        private Element sourceElement;
        public Element SourceElement
        {
            get { return sourceElement; }
            set
            {
                sourceElement = value;
                OnPropertyChanged(nameof(SourceElement));
            }
        }
        private bool isSourceNotFromLink;
        public bool IsSourceNotFromLink
        {
            get { return isSourceNotFromLink; }
            set { isSourceNotFromLink = value; OnPropertyChanged(nameof(IsSourceNotFromLink)); }
        }
        private Element sourceElementType;
        public Element SourceElementType
        {
            get { return sourceElementType; }
            set
            {
                sourceElementType = value;
                OnPropertyChanged(nameof(SourceElementType));
            }
        }
        private string sourceEleFamilyName;
        public string SourceEleFamilyName
        {
            get { return sourceEleFamilyName; }
            set
            {
                sourceEleFamilyName = value;
                OnPropertyChanged(nameof(SourceEleFamilyName));
            }
        }
        private string sourceEleTypeName;
        public string SourceEleTypeName
        {
            get { return sourceEleTypeName; }
            set
            {
                sourceEleTypeName = value;
                OnPropertyChanged(nameof(SourceEleTypeName));
            }
        }
        private ElementId sourceEleId;
        public ElementId SourceEleId
        {
            get { return sourceEleId; }
            set
            {
                sourceEleId = value;
                OnPropertyChanged(nameof(SourceEleId));
            }
        }
        private bool isTargetInstTabEnabled;
        public bool IsTargetInstTabEnabled
        {
            get { return isTargetInstTabEnabled; }
            set
            {
                isTargetInstTabEnabled = value;
                OnPropertyChanged(nameof(IsTargetInstTabEnabled));
            }
        }
        private bool isTargetTypeTabEnabled;
        public bool IsTargetTypeTabEnabled
        {
            get { return isTargetTypeTabEnabled; }
            set
            {
                isTargetTypeTabEnabled = value;
                OnPropertyChanged(nameof(IsTargetTypeTabEnabled));
            }
        }
        private Element targetElement;
        public Element TargetElement
        {
            get { return targetElement; }
            set
            {
                targetElement = value;
                OnPropertyChanged(nameof(TargetElement));
            }
        }
        private Element targetElementType;
        public Element TargetElementType
        {
            get { return targetElementType; }
            set
            {
                targetElementType = value;
                OnPropertyChanged(nameof(TargetElementType));
            }
        }
        private ObservableCollection<Element> sameFamilyTargetElements;
        public ObservableCollection<Element> SameFamilyTargetElements
        {
            get { return sameFamilyTargetElements; }
            set
            {
                sameFamilyTargetElements = value;
                OnPropertyChanged(nameof(SameFamilyTargetElements));
            }
        }
        private ObservableCollection<Element> sameCategoryTargetElements;
        public ObservableCollection<Element> SameCategoryTargetElements
        {
            get { return sameCategoryTargetElements; }
            set
            {
                sameCategoryTargetElements = value;
                OnPropertyChanged(nameof(SameCategoryTargetElements));
            }
        }
        private string targetEleFamilyName;
        public string TargetEleFamilyName
        {
            get { return targetEleFamilyName; }
            set
            {
                targetEleFamilyName = value;
                OnPropertyChanged(nameof(TargetEleFamilyName));
            }
        }
        private string targetEleTypeName;
        public string TargetEleTypeName
        {
            get { return targetEleTypeName; }
            set
            {
                targetEleTypeName = value;
                OnPropertyChanged(nameof(TargetEleTypeName));
            }
        }
        private ElementId targetEleId;
        public ElementId TargetEleId
        {
            get { return targetEleId; }
            set
            {
                targetEleId = value;
                OnPropertyChanged(nameof(TargetEleId));
            }
        }
        private ObservableCollection<ParamObj> sourceInstParams { get;set; }
        public ObservableCollection<ParamObj> SourceInstParams
        {
            get { return sourceInstParams; }
            set
            {
                sourceInstParams = value;
                OnPropertyChanged(nameof(SourceInstParams));
            }
        }
        private ObservableCollection<ParamObj> sourceTypeParams { get; set; }
        public ObservableCollection<ParamObj> SourceTypeParams
        {
            get { return sourceTypeParams; }
            set
            {
                sourceTypeParams = value;
                OnPropertyChanged(nameof(SourceTypeParams));
            }
        }
        private ParamObj selectedSourceParam { get; set; }
        public ParamObj SelectedSourceParam
        {
            get { return selectedSourceParam; }
            set
            {
                selectedSourceParam = value;
                OnPropertyChanged(nameof(SelectedSourceParam));
                CanCopy = true;
            }
        }
        private ObservableCollection<ParamObj> targetInstParams { get; set; }
        public ObservableCollection<ParamObj> TargetInstParams
        {
            get { return targetInstParams; }
            set
            {
                targetInstParams = value;
                OnPropertyChanged(nameof(TargetInstParams));
            }
        }
        private ObservableCollection<ParamObj> targetTypeParams { get; set; }
        public ObservableCollection<ParamObj> TargetTypeParams
        {
            get { return targetTypeParams; }
            set
            {
                targetTypeParams = value;
                OnPropertyChanged(nameof(TargetTypeParams));
            }
        }
        private ParamObj selectedTargetParam { get; set; }
        public ParamObj SelectedTargetParam
        {
            get { return selectedTargetParam; }
            set
            {
                selectedTargetParam = value;
                OnPropertyChanged(nameof(SelectedTargetParam));
                CanCopy = true;
            }
        }
        private bool canCopy { get; set; }
        public bool CanCopy 
        { 
            get { return canCopy; } 
            set { canCopy = value; OnPropertyChanged(nameof(CanCopy)); }
        }
        private dynamic sourceValue { get; set; }
        public dynamic SourceValue
        {
            get { return sourceValue; }
            set
            {
                sourceValue = value;
                OnPropertyChanged(nameof(SourceValue));
            }
        }
        private bool isRounded { get; set; }
        public bool IsRounded 
        { 
            get { return isRounded; } 
            set 
            { 
                isRounded = value; 
                OnPropertyChanged(nameof(IsRounded)); 
            }
        }
        private int roundValue { get; set; }
        public int RoundValue
        {
            get { return roundValue; }
            set { roundValue = value; OnPropertyChanged(nameof(RoundValue)); }
        }
        public ICommand SelectSourceICommand { get; set; }
        public void _SelectSourceICommand(object obj)
        {
            Selection targetSelect = UiDoc.Selection;
            Reference targetSelectRef = null;

            // Catch error if user abort selection
            try
            {
                targetSelectRef = targetSelect.PickObject(ObjectType.Element, "ソースインスタンス選択");
            }
            catch (Exception)
            {
                // TaskDialog.Show("エラー", "参照元を1つ選択してください。");
                MessageBox.Show("参照元を1つ選択してください。", "エラー");
                return;
            }
            if (targetSelectRef == null)
            {
                // TaskDialog.Show("エラー", "参照元を1つ選択してください。");
                MessageBox.Show("参照元を1つ選択してください。", "エラー");
                return;
            }
            SourceElement = ThisDoc.GetElement(targetSelectRef);
            DefineSourceElementInfo(SourceElement);
            IsSourceNotFromLink = true;
            if (TargetElement == null) 
            {
                TargetElement = SourceElement;
                UpdateTarget(TargetElement);
            } 
            /**
            ElementId sourceElementTypeId = SourceElement.GetTypeId();
            SourceElementType = ThisDoc.GetElement(sourceElementTypeId);
            SourceEleFamilyName = SourceElement.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
            SourceEleTypeName = SourceElement.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();
            SourceEleId =  SourceElement.Id;
            ParameterSet targetParamSet = SourceElement.Parameters;
            ObservableCollection<ParamObj> sips = new ObservableCollection<ParamObj>();
            foreach (Parameter targetParam in targetParamSet)
            {
                ParamObj targetParamObj = new ParamObj(targetParam);
                sips.Add(targetParamObj);
            }
            SourceInstParams = new ObservableCollection<ParamObj>();
            sips.OrderBy(p => p.ParamName).ToList().ForEach(p => SourceInstParams.Add(p));

            try
            {
                //var sourceEleTypeId = SourceElement.GetTypeId();
                //var sourceEleType = ThisDoc.GetElement(sourceEleTypeId);
                ParameterSet sourceTypeParamSet = SourceElementType.Parameters;
                ObservableCollection<ParamObj> stps = new ObservableCollection<ParamObj>();
                foreach (Parameter sourceTypeParam in sourceTypeParamSet)
                {
                    ParamObj sourceTypeParamObj = new ParamObj(sourceTypeParam);
                    stps.Add(sourceTypeParamObj);
                }
                SourceTypeParams = new ObservableCollection<ParamObj>();
                stps.OrderBy(p => p.ParamName).ToList().ForEach(p => SourceTypeParams.Add(p));
            }
            catch
            {
            }
            **/
        }
        public ICommand SelectLinkedSourceICommand { get; set; }
        public void _SelectLinkedSourceICommand(object obj)
        {
            Selection targetSelect = UiDoc.Selection;
            Reference targetSelectRef = null;

            // Catch error if user abort selection
            try
            {
                targetSelectRef = targetSelect.PickObject(ObjectType.LinkedElement, "参照元選択");
            }
            catch (Exception)
            {
                // TaskDialog.Show("エラー", "参照元を1つ選択してください。");
                MessageBox.Show("参照元を1つ選択してください。", "エラー");
                return;
            }
            if (targetSelectRef == null)
            {
                // TaskDialog.Show("エラー", "参照元を1つ選択してください。");
                MessageBox.Show("参照元を1つ選択してください。", "エラー");
                return;
            }

            RevitLinkInstance linkedInstance = ThisDoc.GetElement(targetSelectRef) as RevitLinkInstance;
            Document linkedDoc = linkedInstance.GetLinkDocument();
            SourceElement = linkedDoc.GetElement(targetSelectRef.LinkedElementId);
            DefineSourceElementInfo(SourceElement);
            IsSourceNotFromLink = false;
            /**
            ElementId sourceElementTypeId = SourceElement.GetTypeId();
            SourceElementType = ThisDoc.GetElement(sourceElementTypeId);
            SourceEleFamilyName = SourceElement.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
            SourceEleTypeName = SourceElement.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();
            SourceEleId = SourceElement.Id;
            ParameterSet targetParamSet = SourceElement.Parameters;
            ObservableCollection<ParamObj> sips = new ObservableCollection<ParamObj>();
            foreach (Parameter targetParam in targetParamSet)
            {
                ParamObj targetParamObj = new ParamObj(targetParam);
                sips.Add(targetParamObj);
            }
            SourceInstParams = new ObservableCollection<ParamObj>();
            sips.OrderBy(p => p.ParamName).ToList().ForEach(p => SourceInstParams.Add(p));

            try
            {
                //var sourceEleTypeId = SourceElement.GetTypeId();
                //var sourceEleType = ThisDoc.GetElement(sourceEleTypeId);
                ParameterSet sourceTypeParamSet = SourceElementType.Parameters;
                ObservableCollection<ParamObj> stps = new ObservableCollection<ParamObj>();
                foreach (Parameter sourceTypeParam in sourceTypeParamSet)
                {
                    ParamObj sourceTypeParamObj = new ParamObj(sourceTypeParam);
                    stps.Add(sourceTypeParamObj);
                }
                SourceTypeParams = new ObservableCollection<ParamObj>();
                stps.OrderBy(p => p.ParamName).ToList().ForEach(p => SourceTypeParams.Add(p));
            }
            catch
            {
            }
            **/
        }
        private void DefineSourceElementInfo(Element sourceElement)
        {
            if (sourceElement == null) return;
            ElementId sourceElementTypeId = sourceElement.GetTypeId();
            SourceElementType = ThisDoc.GetElement(sourceElementTypeId);
            SourceEleFamilyName = sourceElement.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
            SourceEleTypeName = sourceElement.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();
            SourceEleId = sourceElement.Id;
            ParameterSet targetParamSet = sourceElement.Parameters;
            ObservableCollection<ParamObj> sips = new ObservableCollection<ParamObj>();
            foreach (Parameter targetParam in targetParamSet)
            {
                ParamObj targetParamObj = new ParamObj(targetParam);
                sips.Add(targetParamObj);
            }
            SourceInstParams = new ObservableCollection<ParamObj>();
            sips.OrderBy(p => p.ParamName).ToList().ForEach(p => SourceInstParams.Add(p));

            try
            {
                //var sourceEleTypeId = SourceElement.GetTypeId();
                //var sourceEleType = ThisDoc.GetElement(sourceEleTypeId);
                ParameterSet sourceTypeParamSet = SourceElementType.Parameters;
                ObservableCollection<ParamObj> stps = new ObservableCollection<ParamObj>();
                foreach (Parameter sourceTypeParam in sourceTypeParamSet)
                {
                    ParamObj sourceTypeParamObj = new ParamObj(sourceTypeParam);
                    stps.Add(sourceTypeParamObj);
                }
                SourceTypeParams = new ObservableCollection<ParamObj>();
                stps.OrderBy(p => p.ParamName).ToList().ForEach(p => SourceTypeParams.Add(p));
            }
            catch
            {
            }
        }
        public ICommand SetTargetSameAsSourceICommand { get; set; }
        public void _SetTargetSameAsSourceICommand(object obj)
        {
            TargetElement = SourceElement;
            UpdateTarget(TargetElement);
        }
        public ICommand SelectTargetICommand { get; set; }
        public void _SelectTargetICommand(object obj)
        {
            Selection targetSelect = UiDoc.Selection;
            Reference targetSelectRef = null;

            // Catch error if user abort selection
            try
            {
                targetSelectRef = targetSelect.PickObject(ObjectType.Element, "参照先選択");
            }
            catch (Exception)
            {
                // TaskDialog.Show("エラー", "参照先を1つ選択してください。");
                MessageBox.Show("参照先を1つ選択してください。", "エラー");
                return;
            }
            if (targetSelectRef == null)
            {
                // TaskDialog.Show("エラー", "参照先を1つ選択してください。");
                MessageBox.Show("参照先を1つ選択してください。", "エラー");
                return;
            }
            TargetElement = ThisDoc.GetElement(targetSelectRef);
            UpdateTarget(TargetElement);
        }
        public void UpdateTarget(Element element)
        {
            TargetElement = element;
            if (element == null) return;
            ElementId targetElementTypeId = element.GetTypeId();
            TargetElementType = ThisDoc.GetElement(targetElementTypeId);
            TargetEleFamilyName = TargetElement.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
            TargetEleTypeName = TargetElement.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();
            TargetEleId = TargetElement.Id;

            SameCategoryTargetElements = new ObservableCollection<Element>();

            var allElements = new FilteredElementCollector(ThisDoc)
                .WhereElementIsNotElementType()
                .ToList();

            SameCategoryTargetElements = new ObservableCollection<Element>();

            foreach (var ele in allElements)
            {
                if (ele != null && ele.Category != null
                    && ele.Category.Id.IntegerValue == TargetElement.Category.Id.IntegerValue)
                {
                    SameCategoryTargetElements.Add(ele);
                }
            }

            SameFamilyTargetElements = new ObservableCollection<Element>();

            SameCategoryTargetElements
                .Where(e => e.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                        == TargetEleFamilyName)
                .ToList()
                .ForEach(e => SameFamilyTargetElements.Add(e));

            ParameterSet targetParamSet = TargetElement.Parameters;
            ObservableCollection<ParamObj> tips = new ObservableCollection<ParamObj>();
            foreach (Parameter targetParam in targetParamSet)
            {
                ParamObj targetParamObj = new ParamObj(targetParam);
                tips.Add(targetParamObj);
            }

            var tempSelectedTargetParam = SelectedTargetParam;

            TargetInstParams = new ObservableCollection<ParamObj>();
            tips.OrderBy(p => p.ParamName).ToList().ForEach(p => TargetInstParams.Add(p));

            try
            {
                //var targetEleTypeId = TargetElement.GetTypeId();
                //var targetEleType = ThisDoc.GetElement(targetEleTypeId);
                ParameterSet targetTypeParamSet = TargetElementType.Parameters;
                ObservableCollection<ParamObj> ttps = new ObservableCollection<ParamObj>();
                foreach (Parameter targetTypeParam in targetTypeParamSet)
                {
                    ParamObj targetTypeParamObj = new ParamObj(targetTypeParam);
                    ttps.Add(targetTypeParamObj);
                }
                TargetTypeParams = new ObservableCollection<ParamObj>();
                ttps.OrderBy(p => p.ParamName).ToList().ForEach(p => TargetTypeParams.Add(p));
            }
            catch
            {
            }
            SelectedTargetParam = tempSelectedTargetParam;
        }
        public void CopyParam(Element element)
        {
            if (element == null) return;
            

            StorageType sourceParamStorageType = SelectedSourceParam.Param.StorageType;
            StorageType targetParamStorageType = SelectedTargetParam.Param.StorageType;
            ForgeTypeId sourceParamUnitTypeId = sourceParamStorageType == StorageType.Double ? SelectedSourceParam.Param.GetUnitTypeId() : null;
            ForgeTypeId targetParamUnitTypeId = targetParamStorageType == StorageType.Double ? SelectedTargetParam.Param.GetUnitTypeId() : null;

            if (sourceParamStorageType == targetParamStorageType)
            {
                SourceValue = SelectedSourceParam.ParamValue;
                if (sourceParamStorageType == StorageType.Double && targetParamStorageType == StorageType.Double && sourceParamUnitTypeId != targetParamUnitTypeId)
                {
                    if (sourceParamUnitTypeId != null)
                        SourceValue = Autodesk.Revit.DB.UnitUtils.ConvertFromInternalUnits(SourceValue, sourceParamUnitTypeId);
                    SourceValue = IsRounded ? Math.Round(SourceValue, RoundValue) : SourceValue;
                    if (targetParamUnitTypeId != null)
                        SourceValue = Autodesk.Revit.DB.UnitUtils.ConvertToInternalUnits(SourceValue, targetParamUnitTypeId);
                }
            }
            else
            {
                SourceValue = (sourceParamStorageType == StorageType.String || targetParamStorageType == StorageType.String) ? SelectedSourceParam.ParamValueString : SelectedSourceParam.ParamValue;
                if (sourceParamStorageType != StorageType.String
                    && sourceParamStorageType != StorageType.Integer
                    && sourceParamStorageType != StorageType.Double
                    && (targetParamStorageType == StorageType.Integer
                        || targetParamStorageType == StorageType.Double)
                    || (SelectedSourceParam.Param.Definition.GetDataType().TypeId == "autodesk.spec:spec.bool-1.0.0" 
                        && SelectedTargetParam.Param.Definition.GetDataType().TypeId != "autodesk.spec:spec.bool-1.0.0"
                        && targetParamStorageType != StorageType.String))
                {
                    CanCopy = false;
                    MessageBox.Show("このパラメータをコピーできません。");
                    return;
                }
                if (sourceParamStorageType == StorageType.String && targetParamStorageType == StorageType.Integer)
                {
                    if (SelectedTargetParam.Param.Definition.GetDataType().TypeId == "autodesk.spec:spec.bool-1.0.0")
                    {
                        if (SourceValue.ToString() == "Yes" || SourceValue.ToString() == "はい") SourceValue = 1;
                        else if (SourceValue.ToString() == "No" || SourceValue.ToString() == "いいえ") SourceValue = 0;
                    }
                    else
                    {
                        try
                        {
                            SourceValue = int.Parse(SelectedSourceParam.ParamValue);

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
                if (sourceParamStorageType == StorageType.String && targetParamStorageType == StorageType.Double)
                {
                    try
                    {
                        SourceValue = double.Parse(SelectedSourceParam.ParamValue);
                        SourceValue = Autodesk.Revit.DB.UnitUtils.ConvertToInternalUnits(SourceValue, targetParamUnitTypeId);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                if (sourceParamStorageType == StorageType.Integer && targetParamStorageType == StorageType.Double)
                {
                    SourceValue = Convert.ToDouble(SourceValue);

                    if (sourceParamUnitTypeId != null)
                        SourceValue = Autodesk.Revit.DB.UnitUtils.ConvertFromInternalUnits(SourceValue, sourceParamUnitTypeId);
                    SourceValue = IsRounded ? Math.Round(SourceValue, RoundValue) : SourceValue;
                    if (targetParamUnitTypeId != null)
                        SourceValue = Autodesk.Revit.DB.UnitUtils.ConvertToInternalUnits(SourceValue, targetParamUnitTypeId);
                }
                if (sourceParamStorageType == StorageType.Double && targetParamStorageType == StorageType.Integer
                    && SelectedTargetParam.Param.Definition.GetDataType().TypeId != "autodesk.spec:spec.bool-1.0.0")
                {
                    if (sourceParamUnitTypeId != null)
                        SourceValue = Autodesk.Revit.DB.UnitUtils.ConvertFromInternalUnits(SourceValue, sourceParamUnitTypeId);
                    SourceValue = IsRounded ? Math.Round(SourceValue, RoundValue) : SourceValue;
                    if (targetParamUnitTypeId != null)
                        SourceValue = Autodesk.Revit.DB.UnitUtils.ConvertToInternalUnits(SourceValue, targetParamUnitTypeId);

                    SourceValue = Convert.ToInt32(SourceValue);
                }
                /**
                if (sourceParamStorageType == StorageType.Double && SelectedTargetParam.Param.Definition.GetDataType().TypeId == "autodesk.spec:spec.bool-1.0.0")
                {
                    if (SourceValue.ToString() == "1")
                    {
                        SourceValue = 1;
                        SelectedTargetParam.ParamValueString = BoolYes;
                    }
                    else if (SourceValue.ToString() == "0")
                    {
                        SourceValue = 0;
                        SelectedTargetParam.ParamValueString = BoolNo;
                    }
                    else
                    {
                        MessageBox.Show("このパラメータを設定できません。");
                        return;
                    }
                }
                **/
                //if (targetParamStorageType == StorageType.String)
                //{
                //    SourceValue = SelectedSourceParam.ParamValueString;
                //}
                
            }
            // SelectedTargetParam.ParamValueString = SourceValue.ToString();

            if (SelectedTargetParam.Param.Definition.GetDataType().TypeId == "autodesk.spec:spec.bool-1.0.0")
            {
                if (SourceValue.ToString() == "1")
                {
                    SourceValue = 1;
                    SelectedTargetParam.ParamValueString = BoolYes;
                }
                else if (SourceValue.ToString() == "0")
                {
                    SourceValue = 0;
                    SelectedTargetParam.ParamValueString = BoolNo;
                }
                else
                {
                    MessageBox.Show("このパラメータを設定できません。");
                    return;
                }
            }

            if (SelectedTargetParam.Param.Definition.GetDataType().TypeId != "autodesk.spec:spec.bool-1.0.0")
                SelectedTargetParam.ParamValueString = SelectedSourceParam.ParamValueString;
            bool doesValueContainsLetter = SelectedSourceParam.ParamValueString.Any(c => char.IsLetter(c)) && SelectedTargetParam.Param.Definition.GetDataType().TypeId != "autodesk.spec:spec.bool-1.0.0";
            bool isValueANumber = (sourceParamStorageType == StorageType.Integer || sourceParamStorageType == StorageType.Double) ? true : false;
            if (isValueANumber && doesValueContainsLetter && targetParamStorageType != StorageType.String)
            {
                var spaceIndex = SelectedTargetParam.ParamValueString.IndexOf(" ");
                SelectedTargetParam.ParamValueString = SelectedTargetParam.ParamValueString.Substring(0, spaceIndex + 1);
            }

            SelectedTargetParam.WasUpdated = true;

        }

        public ICommand RefreshICommand { get; set; }
        public void _RefreshICommand(object obj)
        {
            if (SourceElement != null) DefineSourceElementInfo(SourceElement);
            if (TargetElement != null) UpdateTarget(TargetElement);
            MessageBox.Show("更新が完了しました。");
        }
        public ICommand InstanceCopyICommand { get; set; }
        public void _InstanceCopyICommand(object obj)
        {
            Element element = null;

            if (IsTargetInstTabEnabled) element = TargetElement;
            if (IsTargetTypeTabEnabled) element = TargetElementType;
            CopyParam(element);
        }
        public ICommand FamilyCopyICommand { get; set; }
        public void _FamilyCopyICommand(object obj)
        {

            if (SameFamilyTargetElements.Count <= 0) return;

            foreach (var ele in SameFamilyTargetElements)
            {
                Element element = null;
                if (IsTargetInstTabEnabled)
                    element = ele;
                if (IsTargetTypeTabEnabled)
                    element = ele.GetElementType(ThisDoc);
                if (CanCopy) CopyParam(element); else return;
            }

            //UpdateTarget(TargetElement);
        }
        public ICommand CategoryCopyICommand { get; set; }
        public void _CategoryCopyICommand(object obj)
        {

            if (SameCategoryTargetElements.Count <= 0) return;

            foreach (var ele in SameCategoryTargetElements)
            {
                Element element = null;
                if (IsTargetInstTabEnabled)
                    element = ele;
                if (IsTargetTypeTabEnabled)
                    element = ele.GetElementType(ThisDoc);
                if (CanCopy) CopyParam(element); else return;
            }

            //UpdateTarget(TargetElement);
        }
        public ICommand AllEleCopyICommand { get; set; }
        public void _AllEleCopyICommand(object obj)
        {
            var allEle = new FilteredElementCollector(ThisDoc)
                .WhereElementIsNotElementType()
                .ToList();

            foreach (var ele in allEle)
            {
                Element element = null;
                if (IsTargetInstTabEnabled)
                    element = ele;
                if (IsTargetTypeTabEnabled)
                    element = ele.GetElementType(ThisDoc);
                if (CanCopy) CopyParam(element); else return;
            }

            //UpdateTarget(TargetElement);
        }
    }
    public class ParamObj : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public ParamObj(Parameter parameter)
        {
            Param = parameter;
            // ParamName = parameter.Definition.Name;
            // ParamValue = parameter.GetParameterValue();
            // ParamValueString = parameter.AsValueString();
            // IsNotModifiable = !parameter.UserModifiable;
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
                ParamValue = param.GetParameterValue();
                ParamValueString = param.AsValueString();
                IsModifiable = !param.IsReadOnly;
                WasUpdated = false;
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
        private string paramValueString;
        public string ParamValueString
        {
            get { return paramValueString; }
            set 
            { 
                paramValueString = value; 
                OnPropertyChanged(nameof(ParamValueString)); 
            }
        }
        private bool isModifiable;
        public bool IsModifiable
        {
            get { return isModifiable; }
            set { isModifiable = value; OnPropertyChanged(nameof(IsModifiable)); }
        }
        private bool wasUpdated;
        public bool WasUpdated
        {
            get { return wasUpdated; }
            set { wasUpdated = value; OnPropertyChanged(nameof(WasUpdated)); }
        }
        protected void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null) // if there is any subscribers 
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
