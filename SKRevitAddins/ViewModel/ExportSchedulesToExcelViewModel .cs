using System.Collections.ObjectModel;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace SKRevitAddins.ViewModel
{
    public class ExportSchedulesToExcelViewModel : ViewModelBase
    {
        private UIApplication _uiApp;
        private Document _doc;

        public ExportSchedulesToExcelViewModel(UIApplication uiApp)
        {
            _uiApp = uiApp;
            _doc = _uiApp.ActiveUIDocument.Document;
            AvailableSchedules = new ObservableCollection<ScheduleItem>();
            SelectedSchedules = new ObservableCollection<ScheduleItem>();
            ExcelPreviewData = new ObservableCollection<List<string>>();

            LoadSchedules();
        }

        private void LoadSchedules()
        {
            var collector = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => !vs.IsTemplate);

            foreach (var vs in collector)
            {
                AvailableSchedules.Add(new ScheduleItem
                {
                    Schedule = vs,
                    Name = vs.Name
                });
            }
        }

        // Danh sách schedule để hiển thị trong cửa sổ chính
        public ObservableCollection<ScheduleItem> AvailableSchedules { get; set; }

        // Danh sách schedule được chọn để export
        public ObservableCollection<ScheduleItem> SelectedSchedules { get; set; }

        // Nếu cần hiển thị preview
        public ObservableCollection<List<string>> ExcelPreviewData { get; set; }

        public class ScheduleItem
        {
            public string Name { get; set; }
            public ViewSchedule Schedule { get; set; }

            private bool _isSelected;
            public bool IsSelected
            {
                get => _isSelected;
                set
                {
                    _isSelected = value;
                    // OnPropertyChanged(nameof(IsSelected)); // Nếu bạn muốn UI cập nhật
                }
            }
        }
    }
}
