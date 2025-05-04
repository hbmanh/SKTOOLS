using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using SKRevitAddins.LayoutsToDWG.ViewModel;
using Point = System.Windows.Point;
using SheetItem = SKRevitAddins.LayoutsToDWG.ViewModel.SheetItem;
using WinForms = System.Windows.Forms;

namespace SKRevitAddins.LayoutsToDWG
{
    /// <summary>
    /// Interaction logic for LayoutsToDWGWindow.xaml
    /// </summary>
    public partial class LayoutsToDWGWindow : Window
    {
        readonly LayoutsToDWGViewModel _vm;
        object _dragItem;

        public LayoutsToDWGWindow(UIDocument uiDoc)
        {
            InitializeComponent();
            _vm = new LayoutsToDWGViewModel(uiDoc);
            DataContext = _vm;
        }

        //──────── Buttons ──────────────────────────────────────
        void Cancel_Click(object sender, RoutedEventArgs e) => Close();

        void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            using var dlg = new WinForms.FolderBrowserDialog();
            if (dlg.ShowDialog() == WinForms.DialogResult.OK)
                _vm.ExportPath = dlg.SelectedPath;
        }

        void Export_Click(object sender, RoutedEventArgs e)
        {
            var sheets = _vm.SheetItems.Where(x => x.IsSelected)
                                       .OrderBy(x => x.Order)
                                       .ToList();
            if (!sheets.Any()) { Msg("No sheets selected."); return; }
            if (!Directory.Exists(_vm.ExportPath)) { Msg("Path invalid."); return; }

            try
            {
                var dwgOpt = new FilteredElementCollector(_vm.Doc)
                    .OfClass(typeof(ExportDWGSettings))
                    .Cast<ExportDWGSettings>()
                    .First(x => x.Name == _vm.SelectedExportSetup)
                    .GetDWGExportOptions();

                foreach (var si in sheets)                // xuất từng sheet
                {
                    ViewSheet vs = si.Sheet;
                    string numVal = GetParam(vs, _vm.SheetNumberParam);
                    string nameVal = GetParam(vs, _vm.SheetNameParam);

                    string fileName = $"{_vm.Prefix}-{numVal}_{nameVal}";
                    fileName = SanitizeFileName(fileName);

                    _vm.Doc.Export(_vm.ExportPath, fileName,
                                   new[] { vs.Id }, dwgOpt);
                }

                if (_vm.OpenFolderAfterExport)
                    Process.Start("explorer.exe", _vm.ExportPath);

                Close();
            }
            catch (Exception ex) { Msg(ex.Message); }
        }

        static void Msg(string t) => System.Windows.MessageBox.Show(t);

        //──────── Helper: lấy tham số titleblock ───────────────
        static string GetParam(ViewSheet vs, string paramName)
            => vs?.LookupParameter(paramName)?.AsString() ?? "";

        //──────── Helper: tên file hợp lệ Win ──────────────────
        static string SanitizeFileName(string s)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');
            return s.Trim();
        }

        //──────── Drag‑drop reorder ────────────────────────────
        void SheetDataGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
            => _dragItem = RowItemAt<UIElement>(e.GetPosition(SheetDataGrid));

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

                // cập nhật lại thứ tự
                for (int k = 0; k < _vm.SheetItems.Count; k++)
                    _vm.SheetItems[k].Order = k + 1;
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
