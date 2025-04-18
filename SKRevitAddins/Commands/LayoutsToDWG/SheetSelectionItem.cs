using Autodesk.Revit.DB;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    public class SheetSelectionItem : INotifyPropertyChanged
    {
        public ViewSheet Sheet { get; set; }

        private string _sheetName;
        public string SheetName
        {
            get => _sheetName;
            set { _sheetName = value; On(); }
        }

        private string _level;
        public string Level
        {
            get => _level;
            set { _level = value; On(); }
        }

        private string _paperSize;
        public string PaperSize
        {
            get => _paperSize;
            set { _paperSize = value; On(); }
        }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; On(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        void On([CallerMemberName] string p = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
    }
}
