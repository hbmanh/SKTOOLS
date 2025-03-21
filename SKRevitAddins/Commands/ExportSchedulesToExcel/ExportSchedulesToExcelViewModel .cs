using System.Collections.ObjectModel;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace SKRevitAddins.ExportSchedulesToExcel
{
    public class ExportSchedulesToExcelViewModel : INotifyPropertyChanged
    {
        private UIApplication _uiApp;
        private Document _hostDoc;

        public ExportSchedulesToExcelViewModel(UIApplication uiApp)
        {
            _uiApp = uiApp;
            _hostDoc = _uiApp.ActiveUIDocument.Document;

            Documents = new ObservableCollection<DocumentItem>();
            FilteredDocuments = new ObservableCollection<DocumentItem>();
            AllSchedules = new ObservableCollection<ScheduleItem>();
            FilteredSchedules = new ObservableCollection<ScheduleItem>();
            SelectedSchedules = new ObservableCollection<ScheduleItem>();

            LoadDocuments();
            foreach (var d in Documents)
                FilteredDocuments.Add(d);

            ExportStatusMessage = "";
            ExportProgressValue = 0;
        }

        // Thông báo trạng thái
        private string _exportStatusMessage;
        public string ExportStatusMessage
        {
            get => _exportStatusMessage;
            set
            {
                _exportStatusMessage = value;
                OnPropertyChanged();
            }
        }

        // Tiến độ xuất (0–100), liên kết với ProgressBar
        private int _exportProgressValue;
        public int ExportProgressValue
        {
            get => _exportProgressValue;
            set
            {
                _exportProgressValue = value;
                OnPropertyChanged();
            }
        }

        // Nếu muốn hủy (Cancel) thì dùng thuộc tính này
        private bool _isCancelled;
        public bool IsCancelled
        {
            get => _isCancelled;
            set { _isCancelled = value; OnPropertyChanged(); }
        }

        // Danh sách Document
        public ObservableCollection<DocumentItem> Documents { get; set; }
        public ObservableCollection<DocumentItem> FilteredDocuments { get; set; }

        private DocumentItem _selectedDocumentItem;
        public DocumentItem SelectedDocumentItem
        {
            get => _selectedDocumentItem;
            set
            {
                _selectedDocumentItem = value;
                OnPropertyChanged();
                LoadSchedulesForSelectedDoc();
            }
        }

        // Danh sách Schedule
        public ObservableCollection<ScheduleItem> AllSchedules { get; set; }
        public ObservableCollection<ScheduleItem> FilteredSchedules { get; set; }
        public ObservableCollection<ScheduleItem> SelectedSchedules { get; set; }

        private void LoadDocuments()
        {
            // Host Document
            Documents.Add(new DocumentItem
            {
                DisplayName = "Host Document",
                DocRef = _hostDoc,
                IsLink = false
            });

            // Link
            var linkInsts = new FilteredElementCollector(_hostDoc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            foreach (var li in linkInsts)
            {
                var linkDoc = li.GetLinkDocument();
                if (linkDoc != null)
                {
                    Documents.Add(new DocumentItem
                    {
                        DisplayName = "Link: " + linkDoc.Title,
                        DocRef = linkDoc,
                        IsLink = true
                    });
                }
            }

            SelectedDocumentItem = Documents.FirstOrDefault();
        }

        private void LoadSchedulesForSelectedDoc()
        {
            AllSchedules.Clear();
            FilteredSchedules.Clear();
            SelectedSchedules.Clear();

            if (SelectedDocumentItem?.DocRef == null) return;
            var doc = SelectedDocumentItem.DocRef;

            var vsCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => !vs.IsTemplate);

            foreach (var vs in vsCollector)
            {
                AllSchedules.Add(new ScheduleItem
                {
                    Name = vs.Name,
                    Schedule = vs
                });
            }

            foreach (var s in AllSchedules)
                FilteredSchedules.Add(s);
        }

        public void FilterDocumentByKeyword(string keyword)
        {
            FilteredDocuments.Clear();
            if (string.IsNullOrWhiteSpace(keyword))
            {
                foreach (var d in Documents)
                    FilteredDocuments.Add(d);
                return;
            }
            keyword = keyword.ToLower();
            var fil = Documents.Where(d => d.DisplayName.ToLower().Contains(keyword)).ToList();
            foreach (var docItem in fil)
                FilteredDocuments.Add(docItem);
        }

        public void FilterScheduleByKeyword(string keyword)
        {
            FilteredSchedules.Clear();
            if (string.IsNullOrWhiteSpace(keyword))
            {
                foreach (var s in AllSchedules)
                    FilteredSchedules.Add(s);
                return;
            }
            keyword = keyword.ToLower();
            var f = AllSchedules.Where(s => s.Name.ToLower().Contains(keyword)).ToList();
            foreach (var item in f)
                FilteredSchedules.Add(item);
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string prop = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        // DocumentItem
        public class DocumentItem
        {
            public string DisplayName { get; set; }
            public Document DocRef { get; set; }
            public bool IsLink { get; set; }
        }

        // ScheduleItem
        public class ScheduleItem : INotifyPropertyChanged
        {
            private bool _isSelected;
            public bool IsSelected
            {
                get => _isSelected;
                set { _isSelected = value; OnPropertyChanged(); }
            }

            public string Name { get; set; }
            public ViewSchedule Schedule { get; set; }

            public event PropertyChangedEventHandler PropertyChanged;
            private void OnPropertyChanged([CallerMemberName] string prop = null)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
            }
        }
    }
}
