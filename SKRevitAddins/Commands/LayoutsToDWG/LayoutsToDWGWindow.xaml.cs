// SKRevitAddins.Commands.LayoutsToDWG.LayoutsToDWGWindow.xaml.cs
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using SKRevitAddins.Commands.LayoutsToDWG.Helpers;
using SKRevitAddins.Commands.LayoutsToDWG.ViewModel;
using SKRevitAddins.Commands.DWGExport;
using Point = System.Windows.Point;
using SheetItem = SKRevitAddins.Commands.LayoutsToDWG.ViewModel.SheetItem;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    public partial class LayoutsToDWGWindow : Window
    {
        private readonly LayoutsToDWGViewModel _vm;
        private object _dragItem;

        public LayoutsToDWGWindow(UIDocument uiDoc)
        {
            InitializeComponent();
            _vm = new LayoutsToDWGViewModel(uiDoc);
            DataContext = _vm;
        }

        private void Cancel_Click(object s, RoutedEventArgs e) => Close();

        private void BrowseFolder_Click(object s, RoutedEventArgs e)
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                _vm.ExportPath = dlg.SelectedPath;
        }

        private void Export_Click(object s, RoutedEventArgs e)
        {
            var sheets = _vm.SheetItems.Where(x => x.IsSelected).OrderBy(x => x.Order).ToList();
            if (!sheets.Any()) { Msg("No sheets selected."); return; }
            if (!Directory.Exists(_vm.ExportPath)) { Msg("Path invalid."); return; }

            // 1. DWG export
            var dwgOpt = new FilteredElementCollector(_vm.Doc)
                .OfClass(typeof(ExportDWGSettings))
                .Cast<ExportDWGSettings>()
                .First(x => x.Name == _vm.SelectedExportSetup)
                .GetDWGExportOptions();
            dwgOpt.MergedViews = _vm.MergeSheets;

            var ids = sheets.Select(sht => sht.SheetId).ToList();

            // Đường dẫn xuất trực tiếp
            if (!_vm.MergeSheets)
            {
                _vm.Doc.Export(_vm.ExportPath, "", ids, dwgOpt);
                AfterExport(); return;
            }

            // 2. Chế độ merge
            string temp = Path.Combine(_vm.ExportPath, "TempDWG");
            Directory.CreateDirectory(temp);
            _vm.Doc.Export(temp, "", ids, dwgOpt);

            string json = Path.Combine(_vm.ExportPath, "merge_command.json");
            string script = Path.Combine(_vm.ExportPath, "run_merge.scr");
            string[] dwgs = Directory.GetFiles(temp, "*.dwg");
            string outDwg = Path.Combine(_vm.ExportPath, _vm.MergeFilename + ".dwg");

            DWGMergeHelper.GenerateMergeCommandJson(json, dwgs, outDwg, true, _vm.OpenFolderAfterExport);
            DWGMergeHelper.GenerateMergeCommandScript(script);

            string acad = AutoCADLocator.FindLowestAcadExe();
            if (string.IsNullOrEmpty(acad)) { Msg("AutoCAD not found."); return; }

            DWGMergeHelper.RunAutoCADMerge(acad, script);
            AfterExport();
        }

        private void AfterExport()
        {
            if (_vm.OpenFolderAfterExport) Process.Start(_vm.ExportPath);
            Close();
        }
        private static void Msg(string txt) => MessageBox.Show(txt);

        // ────────── Drag‑and‑drop re‑order ──────────
        private void SheetDataGrid_PreviewMouseLeftButtonDown(object s, MouseButtonEventArgs e)
            => _dragItem = RowItemAt<UIElement>(e.GetPosition(SheetDataGrid));

        private void SheetDataGrid_MouseMove(object s, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && _dragItem != null)
                DragDrop.DoDragDrop(SheetDataGrid, _dragItem, DragDropEffects.Move);
        }
        private void SheetDataGrid_Drop(object s, DragEventArgs e)
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
        private T RowItemAt<T>(Point p) where T : class
        {
            var el = SheetDataGrid.InputHitTest(p) as DependencyObject;
            while (el != null && el is not DataGridRow) el = VisualTreeHelper.GetParent(el);
            return (el as DataGridRow)?.Item as T;
        }
    }
}
