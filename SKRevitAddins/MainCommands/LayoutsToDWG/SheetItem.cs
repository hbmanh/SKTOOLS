using Autodesk.Revit.DB;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace SKRevitAddins.LayoutsToDWG.ViewModel
{
    public class SheetItem : INotifyPropertyChanged
    {
        public SheetItem(ViewSheet s, string sheetSetName)
        {
            Sheet = s;
            SheetSetName = sheetSetName;
        }

        public ViewSheet Sheet { get; }
        public string SheetNumber => Sheet.SheetNumber;
        public string SheetName => Sheet.Name;
        public string SheetSetName { get; }

        bool _sel = true;
        public bool IsSelected
        {
            get => _sel;
            set { _sel = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }
}
