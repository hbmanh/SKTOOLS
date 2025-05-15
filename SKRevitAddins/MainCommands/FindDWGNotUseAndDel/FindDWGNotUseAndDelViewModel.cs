using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKRevitAddins.ViewModel;

namespace SKRevitAddins.FindDWGNotUseAndDel
{
    public class FindDWGNotUseAndDelViewModel : ViewModelBase
    {
        private UIApplication _uiApp;
        private Document _doc;

        public FindDWGNotUseAndDelViewModel(UIApplication uiApp)
        {
            _uiApp = uiApp;
            _doc = _uiApp.ActiveUIDocument.Document;

            ImportedDWGs = new ObservableCollection<DwgItem>();
            IsDeleteAll = true;

            LoadDwgData();
        }

        private void LoadDwgData()
        {
            var col = new FilteredElementCollector(_doc)
                .OfClass(typeof(ImportInstance))
                .WhereElementIsNotElementType();

            foreach (ImportInstance importInst in col)
            {
                Element typeElem = _doc.GetElement(importInst.GetTypeId());
                if (typeElem == null) continue;

                string fileName = typeElem.Name ?? "";
                string lower = fileName.ToLower();
                //if (!lower.EndsWith(".dwg") && !lower.EndsWith(".dxf"))
                //    continue;

                bool pinned = importInst.Pinned;
                View v = _doc.GetElement(importInst.OwnerViewId) as View;
                Group grp = _doc.GetElement(importInst.GroupId) as Group;

                ImportedDWGs.Add(new DwgItem
                {
                    InstanceId = importInst.Id.IntegerValue.ToString(),
                    TypeId = typeElem.Id.IntegerValue.ToString(),
                    FileName = fileName,
                    InsertType = "Import",
                    Pinned = pinned ? "yes" : "no",
                    OwnerView = (v != null) ? v.Name : "",
                    Group = (grp != null) ? grp.Name : ""
                });
            }
        }

        // --- Thuộc tính chính ---
        private ObservableCollection<DwgItem> _importedDWGs;
        public ObservableCollection<DwgItem> ImportedDWGs
        {
            get => _importedDWGs;
            set { _importedDWGs = value; OnPropertyChanged(nameof(ImportedDWGs)); }
        }

        private bool _isDeleteAll;
        public bool IsDeleteAll
        {
            get => _isDeleteAll;
            set { _isDeleteAll = value; OnPropertyChanged(nameof(IsDeleteAll)); }
        }
        private IList<DwgItem> _selectedDWGs;
        public IList<DwgItem> SelectedDWGs
        {
            get => _selectedDWGs;
            set
            {
                _selectedDWGs = value;
                OnPropertyChanged(nameof(SelectedDWGs));
            }
        }


        // --- Filter / Reload
        public void FilterDWGs(string keyword)
        {
            if (string.IsNullOrWhiteSpace(keyword))
            {
                // load lại
                ImportedDWGs.Clear();
                LoadDwgData();
                return;
            }

            keyword = keyword.ToLower();
            var filtered = ImportedDWGs
                .Where(d => d.FileName.ToLower().Contains(keyword)
                         || d.OwnerView.ToLower().Contains(keyword)
                         || d.InstanceId.Contains(keyword)
                         || d.TypeId.Contains(keyword))
                .ToList();

            ImportedDWGs.Clear();
            foreach (var f in filtered)
            {
                ImportedDWGs.Add(f);
            }
        }


        // DwgItem class
        public class DwgItem
        {
            public string InstanceId { get; set; }
            public string TypeId { get; set; }
            public string FileName { get; set; }
            public string InsertType { get; set; }
            public string Pinned { get; set; }
            public string OwnerView { get; set; }
            public string Group { get; set; }
        }
    }
}
