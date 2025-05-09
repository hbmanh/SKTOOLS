using Autodesk.Revit.UI;
using SKRevitAddins.LayoutsToDWG.ViewModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Point = System.Windows.Point;

namespace SKRevitAddins.LayoutsToDWG
{
    public partial class LayoutsToDWGWindow : Window
    {
        readonly LayoutsToDWGViewModel _vm;
        object _dragItem;

        public LayoutsToDWGWindow(UIDocument uiDoc)
        {
            InitializeComponent();
            _vm = new LayoutsToDWGViewModel(uiDoc);
            DataContext = _vm;
            _vm.RequestClose = Close;
        }

        private void SetsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _vm.SelectedSheetSets.Clear();
            foreach (string s in SetsListBox.SelectedItems)
                _vm.SelectedSheetSets.Add(s);
        }

        // Phần drag-drop giữ nguyên từ bản trước
        void SheetDataGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
            => _dragItem = RowItemAt<SheetItem>(e.GetPosition(SheetDataGrid));

        void SheetDataGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && _dragItem != null)
                DragDrop.DoDragDrop(SheetDataGrid, _dragItem, DragDropEffects.Move);
        }

        void SheetDataGrid_Drop(object sender, DragEventArgs e)
        {
            var target = RowItemAt<SheetItem>(e.GetPosition(SheetDataGrid));
            if (_dragItem is SheetItem from && target is SheetItem to && from != to)
            {
                int i = _vm.SheetItems.IndexOf(from);
                int j = _vm.SheetItems.IndexOf(to);
                if (i >= 0 && j >= 0) _vm.SheetItems.Move(i, j);
            }
            _dragItem = null;
        }

        T RowItemAt<T>(Point p) where T : class
        {
            var el = SheetDataGrid.InputHitTest(p) as DependencyObject;
            while (el != null && el is not DataGridRow) el = VisualTreeHelper.GetParent(el);
            return (el as DataGridRow)?.Item as T;
        }
    }

}