using Autodesk.Revit.DB;
using SKRevitAddins.LayoutsToDWG.ViewModel;

namespace SKRevitAddins.LayoutsToDWG
{
    public class SheetItem : ViewModelBase
    {
        public SheetItem(ViewSheet s) => (Sheet, SheetNumber, SheetName) =
            (s, s.SheetNumber, s.Name);

        public ViewSheet Sheet { get; }
        public string SheetNumber { get; }
        public string SheetName { get; }

        public int SheetNumSortKey =>
            int.TryParse(SheetNumber, out int n) ? n : int.MaxValue;

        bool _sel = true;
        public bool IsSelected { get => _sel; set { _sel = value; OnPropertyChanged(); } }
    }
}