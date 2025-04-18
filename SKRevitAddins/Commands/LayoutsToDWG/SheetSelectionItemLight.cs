using Autodesk.Revit.DB;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    public class SheetSelectionItemLight : INotifyPropertyChanged
    {
        public ElementId Id { get; set; }

        private string _sheetName;
        public string SheetName
        {
            get => _sheetName;
            set { _sheetName = value; OnPropertyChanged(); }
        }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string prop = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }
}
