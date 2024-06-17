using System;
using System.Collections.Generic;
using Autodesk.Revit.UI;
using System.Linq;
using Autodesk.Revit.DB;
using System.Collections.ObjectModel;
using Application = Autodesk.Revit.ApplicationServices.Application;
using System.Windows.Controls;

namespace SKToolsAddins.ViewModel
{
    public class FindDWGNotUsedAndDelViewModel : ViewModelBase
    {
        private UIApplication UiApp;
        private UIDocument UiDoc;
        private Application ThisApp;
        private Document ThisDoc;

        public FindDWGNotUsedAndDelViewModel(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = UiApp.ActiveUIDocument;
            ThisApp = UiApp.Application;
            ThisDoc = UiDoc.Document;

            ImportedDWGs = new FilteredElementCollector(ThisDoc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(ImportInstance))
                .Cast<ImportInstance>() 
                .Where(x => !x.IsLinked && x.Category.Id.IntegerValue != (int)BuiltInCategory.OST_RasterImages)
                .ToList();
            SelImportedDWG = ImportedDWGs.FirstOrDefault();

            DWGNames = new ObservableCollection<string>();
            DWGIds = new ObservableCollection<int>();
            DWGRefViews = new ObservableCollection<string>();
            UpdateInforOfDWG();

        }

        #region Properties

        private List<ImportInstance> _importedDwGs;
        public List<ImportInstance> ImportedDWGs
        {
            get { return _importedDwGs; }
            set
            {
                _importedDwGs = value;
                OnPropertyChanged(nameof(ImportedDWGs));
            }
        }

        private ImportInstance _selImportedDWG;
        public ImportInstance SelImportedDWG
        {
            get { return _selImportedDWG; }
            set
            {
                _selImportedDWG = value;
                OnPropertyChanged(nameof(SelImportedDWG));
                UpdateInforOfDWG();
            }
        }

        private ObservableCollection<string> _dwgNames;
        public ObservableCollection<string> DWGNames
        {
            get { return _dwgNames; }
            set
            {
                _dwgNames = value;
                OnPropertyChanged(nameof(DWGNames));
            }
        }

        private ObservableCollection<int> _dwgIds;
        public ObservableCollection<int> DWGIds
        {
            get { return _dwgIds; }
            set
            {
                _dwgIds = value;
                OnPropertyChanged(nameof(DWGIds));
            }
        }

        private ObservableCollection<string> _dwgRefViews;

        public ObservableCollection<string> DWGRefViews
        {
            get { return _dwgRefViews; }
            set
            {
                _dwgRefViews = value;
                OnPropertyChanged(nameof(DWGRefViews));
            }
        }

        #endregion


        private void UpdateInforOfDWG()
        {
            List<ImportedDWGInfo> importedCADInfoList = new List<ImportedDWGInfo>();

            foreach (var importInstance in ImportedDWGs)
            {
                var importedCADInfo = new ImportedDWGInfo
                {
                    DWGName = importInstance.Name,
                    //CadId = importInstance.Id.IntegerValue.ToString(),
                    //CadViews = GetImportedViews(importInstance)
                };

                importedCADInfoList.Add(importedCADInfo);
            }

        }

        public class ImportedDWGInfo
        {
            public string DWGName { get; set; }
            public string DWGId { get; set; }
            public List<string> DWGRefViews { get; set; }
        }



    }
}
