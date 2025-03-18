using System.Collections.ObjectModel;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
using System.Data;
using System.Collections.Generic;

namespace SKRevitAddins.ViewModel
{
    public class ExportSchedulesToExcelViewModel : ViewModelBase
    {
        private UIApplication _uiApp;
        private Document _hostDoc;

        public ExportSchedulesToExcelViewModel(UIApplication uiApp)
        {
            _uiApp = uiApp;
            _hostDoc = uiApp.ActiveUIDocument.Document;

            AllDocuments = new ObservableCollection<DocumentItem>();
            FilteredDocuments = new ObservableCollection<DocumentItem>();

            AvailableSchedules = new ObservableCollection<ScheduleItem>();
            FilteredSchedules = new ObservableCollection<ScheduleItem>();

            SelectedSchedules = new ObservableCollection<ScheduleItem>();
            ExcelPreviewTabs = new ObservableCollection<PreviewTab>();

            LoadDocuments();
            // Mặc định hiển thị hết Document
            foreach (var docItem in AllDocuments)
            {
                FilteredDocuments.Add(docItem);
            }
        }

        #region Thuộc tính chính

        // Danh sách Document (host + link)
        public ObservableCollection<DocumentItem> AllDocuments { get; set; }
        public ObservableCollection<DocumentItem> FilteredDocuments { get; set; }

        private DocumentItem _selectedDocumentItem;
        public DocumentItem SelectedDocumentItem
        {
            get => _selectedDocumentItem;
            set
            {
                _selectedDocumentItem = value;
                OnPropertyChanged(nameof(SelectedDocumentItem));
                LoadSchedulesForSelectedDoc();
            }
        }

        // Danh sách schedule
        public ObservableCollection<ScheduleItem> AvailableSchedules { get; set; }
        public ObservableCollection<ScheduleItem> FilteredSchedules { get; set; }
        public ObservableCollection<ScheduleItem> SelectedSchedules { get; set; }

        // Danh sách tab preview
        public ObservableCollection<PreviewTab> ExcelPreviewTabs { get; set; }

        #endregion

        #region Load Documents & Schedules

        private void LoadDocuments()
        {
            // 1) Host
            AllDocuments.Add(new DocumentItem
            {
                DisplayName = "Host Document",
                DocRef = _hostDoc,
                IsLink = false
            });

            // 2) Link
            var linkInstances = new FilteredElementCollector(_hostDoc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            foreach (var linkInst in linkInstances)
            {
                var linkDoc = linkInst.GetLinkDocument();
                if (linkDoc != null)
                {
                    AllDocuments.Add(new DocumentItem
                    {
                        DisplayName = "Link: " + linkDoc.Title,
                        DocRef = linkDoc,
                        IsLink = true
                    });
                }
            }

            // Mặc định chọn Host
            SelectedDocumentItem = AllDocuments.FirstOrDefault();
        }

        private void LoadSchedulesForSelectedDoc()
        {
            AvailableSchedules.Clear();
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
                AvailableSchedules.Add(new ScheduleItem
                {
                    Name = vs.Name,
                    Schedule = vs
                });
            }

            // Hiển thị hết
            foreach (var s in AvailableSchedules)
            {
                FilteredSchedules.Add(s);
            }
        }

        #endregion

        #region Search Document & Schedule

        public void FilterDocumentByKeyword(string keyword)
        {
            if (string.IsNullOrWhiteSpace(keyword))
            {
                FilteredDocuments.Clear();
                foreach (var docItem in AllDocuments)
                    FilteredDocuments.Add(docItem);
                return;
            }
            keyword = keyword.ToLower();

            var filtered = AllDocuments
                .Where(d => d.DisplayName.ToLower().Contains(keyword))
                .ToList();

            FilteredDocuments.Clear();
            foreach (var docItem in filtered)
            {
                FilteredDocuments.Add(docItem);
            }
        }

        public void FilterScheduleByKeyword(string keyword)
        {
            if (string.IsNullOrWhiteSpace(keyword))
            {
                FilteredSchedules.Clear();
                foreach (var s in AvailableSchedules)
                    FilteredSchedules.Add(s);
                return;
            }
            keyword = keyword.ToLower();

            var filtered = AvailableSchedules
                .Where(s => s.Name.ToLower().Contains(keyword))
                .ToList();

            FilteredSchedules.Clear();
            foreach (var item in filtered)
            {
                FilteredSchedules.Add(item);
            }
        }

        #endregion

        #region Preview

        public void LoadPreviewTabsFromSelectedSchedules()
        {
            ExcelPreviewTabs.Clear();

            // Mỗi schedule => 1 tab
            foreach (var sched in SelectedSchedules)
            {
                var dt = new DataTable();
                var tableData = sched.Schedule.GetTableData();
                var section = tableData.GetSectionData(SectionType.Body);
                int rowCount = section.NumberOfRows;
                int colCount = section.NumberOfColumns;

                // Tạo cột
                for (int c = 0; c < colCount; c++)
                {
                    dt.Columns.Add("Column" + c, typeof(string));
                }

                // Tạo row
                for (int r = 0; r < rowCount; r++)
                {
                    var rowVals = new object[colCount];
                    for (int c = 0; c < colCount; c++)
                    {
                        string cellText = sched.Schedule.GetCellText(SectionType.Body, r, c);
                        rowVals[c] = cellText;
                    }
                    dt.Rows.Add(rowVals);
                }

                ExcelPreviewTabs.Add(new PreviewTab
                {
                    SheetName = sched.Name,
                    SheetData = dt
                });
            }
        }

        #endregion

        #region Nested Classes

        // DocumentItem
        public class DocumentItem
        {
            public string DisplayName { get; set; }
            public Document DocRef { get; set; }
            public bool IsLink { get; set; }
        }

        // ScheduleItem
        public class ScheduleItem
        {
            public string Name { get; set; }
            public ViewSchedule Schedule { get; set; }
            public bool IsSelected { get; set; }
        }

        // PreviewTab
        public class PreviewTab
        {
            public string SheetName { get; set; }
            public DataTable SheetData { get; set; }
        }

        #endregion
    }
}
