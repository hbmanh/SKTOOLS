using Autodesk.Revit.UI;
using SKRevitAddins.Utils;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;

namespace SKRevitAddins.SleeveChecker
{
    public class ErrorItem : ViewModelBase
    {
        public string Type { get; set; }
        public string Frame { get; set; }
        public string MEP { get; set; }
        public string Message { get; set; }
        public Autodesk.Revit.DB.ElementId TargetElementId { get; set; }
        public ICommand ShowCommand { get; set; }
    }

    public class SleeveCheckerViewModel : ViewModelBase
    {
        public enum RequestId { None, Preview, Apply, ShowError }
        public RequestId PendingRequest { get; set; } = RequestId.None;

        public double OffsetX { get; set; } = 0.15;
        public double OffsetY { get; set; } = 0.2;
        public double MaxOD { get; set; } = 300;
        public double RatioB { get; set; } = 0.33;
        public double RatioC { get; set; } = 1.25;

        public ObservableCollection<ErrorItem> Errors { get; } = new();
        public ErrorItem SelectedError { get; set; }

        private ExternalEvent _exEvent;

        public ICommand PreviewCommand { get; }
        public ICommand ApplyCommand { get; }
        public ICommand ShowErrorCommand { get; }
        public ICommand ExportCommand { get; }
        public ICommand CancelCommand { get; }

        public SleeveCheckerViewModel(UIApplication uiapp)
        {
            PreviewCommand = new RelayCommand(_ =>
            {
                PendingRequest = RequestId.Preview;
                _exEvent?.Raise();
            });
            ApplyCommand = new RelayCommand(_ =>
            {
                PendingRequest = RequestId.Apply;
                _exEvent?.Raise();
            });
            ShowErrorCommand = new RelayCommand(_ =>
            {
                PendingRequest = RequestId.ShowError;
                _exEvent?.Raise();
            });
            ExportCommand = new RelayCommand(_ => ExportErrors());
            CancelCommand = new RelayCommand(_ => System.Windows.Application.Current.Windows[0]?.Close());
        }

        public void SetExternalEvent(ExternalEvent exEvent) => _exEvent = exEvent;

        public void DoPreview(UIApplication app)
        {
            Errors.Clear();
            using (var tr = new Autodesk.Revit.DB.Transaction(app.ActiveUIDocument.Document, "Preview"))
            {
                tr.Start();
                SleeveCheckerLogic.Run(app.ActiveUIDocument.Document, this, Errors, previewOnly: true, app);
                tr.RollBack();
            }
        }
        public void DoApply(UIApplication app)
        {
            Errors.Clear();
            using (var tr = new Autodesk.Revit.DB.Transaction(app.ActiveUIDocument.Document, "Apply"))
            {
                tr.Start();
                SleeveCheckerLogic.Run(app.ActiveUIDocument.Document, this, Errors, previewOnly: false, app);
                tr.Commit();
            }
        }
        public void DoShowError(UIApplication app)
        {
            if (SelectedError?.TargetElementId != null)
                SleeveCheckerLogic.ShowElement(app, SelectedError.TargetElementId);
        }
        private void ExportErrors()
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV File (*.csv)|*.csv",
                FileName = "SleeveErrors.csv"
            };
            if (dialog.ShowDialog() == true)
            {
                System.IO.File.WriteAllLines(dialog.FileName,
                    new[] { "Type,Frame,MEP,Message" }
                    .Concat(Errors.Select(e => $"{e.Type},{e.Frame},{e.MEP},{e.Message}")));
                System.Windows.MessageBox.Show("Xuất lỗi thành công!");
            }
        }
    }
}
