//using Autodesk.Revit.UI;
//using Microsoft.Win32;
//using System;
//using System.Linq;
//using System.Text;
//using System.Windows;
//using System.Windows.Controls;
//using System.Windows.Media.Imaging;

//namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
//{
//    public partial class PlaceElementsFromBlocksCadWpfWindow : Window
//    {
//        public PlaceElementsFromBlocksCadWpfWindow(
//            ExternalEvent exEvent,
//            AutoPlaceElementFrBlockCADRequestHandler handler,
//            AutoPlaceElementFrBlockCADViewModel viewModel)
//        {
//            InitializeComponent();
//            DataContext = viewModel;

//            viewModel.UpdateStatus = msg =>
//            {
//                Dispatcher.Invoke(() => StatusText.Text = msg);
//            };

//            // Giá trị phổ biến nhất cho batch ComboBox
//            var mostUsedCat = viewModel.BlockMappings
//                .Where(bm => bm.SelectedCategoryMapping != null)
//                .GroupBy(bm => bm.SelectedCategoryMapping)
//                .OrderByDescending(g => g.Count())
//                .Select(g => g.Key)
//                .FirstOrDefault();
//            if (mostUsedCat != null)
//                BatchCategoryComboBox.SelectedItem = mostUsedCat;

//            var mostUsedFam = viewModel.BlockMappings
//                .Where(bm => bm.SelectedFamilyMapping != null)
//                .GroupBy(bm => bm.SelectedFamilyMapping)
//                .OrderByDescending(g => g.Count())
//                .Select(g => g.Key)
//                .FirstOrDefault();
//            if (mostUsedFam != null)
//                BatchFamilyComboBox.SelectedItem = mostUsedFam;

//            var mostUsedType = viewModel.BlockMappings
//                .Where(bm => bm.SelectedTypeSymbolMapping != null)
//                .GroupBy(bm => bm.SelectedTypeSymbolMapping)
//                .OrderByDescending(g => g.Count())
//                .Select(g => g.Key)
//                .FirstOrDefault();
//            if (mostUsedType != null)
//                BatchTypeComboBox.SelectedItem = mostUsedType;

//            // Offset 2600 đã có trong XAML

//            try
//            {
//                string version = viewModel.UiApp.Application.VersionNumber;
//                string iconPath = System.IO.Path.Combine(
//                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
//                    "Autodesk", "Revit", "Addins", version, "SKTools.bundle", "Icon", "shinken.png");

//                if (System.IO.File.Exists(iconPath))
//                {
//                    LogoImage.Source = new BitmapImage(new Uri(iconPath, UriKind.Absolute));
//                }
//            }
//            catch
//            {
//                // Có thể log lỗi ở đây nếu cần
//            }
//        }

//        private void ExportError_Click(object sender, RoutedEventArgs e)
//        {
//            if (DataContext is AutoPlaceElementFrBlockCADViewModel vm)
//            {
//                var failed = vm.BlockMappings
//                    .Where(bm => bm.HasPlacementRun && bm.PlacedCount != bm.BlockCount)
//                    .ToList();

//                if (failed.Count == 0)
//                {
//                    MessageBox.Show("Không có block nào bị lỗi!", "Note", MessageBoxButton.OK, MessageBoxImage.Information);
//                    return;
//                }

//                // Chọn nơi lưu file
//                var sfd = new SaveFileDialog
//                {
//                    Title = "Chọn nơi lưu file xuất lỗi",
//                    Filter = "CSV file (*.csv)|*.csv",
//                    FileName = $"BlockErrorNotes-{DateTime.Now:yyyyMMdd-HHmmss}.csv"
//                };
//                if (sfd.ShowDialog() != true)
//                    return;

//                // Chuẩn bị dữ liệu CSV
//                var csv = new StringBuilder();
//                csv.AppendLine("Block Name,Count,Placed,Failure Note");
//                foreach (var bm in failed)
//                {
//                    string blockName = bm.DisplayBlockName;
//                    string note = string.IsNullOrWhiteSpace(bm.FailureNote)
//                        ? $"Chỉ đặt được {bm.PlacedCount}/{bm.BlockCount} instance."
//                        : bm.FailureNote;
//                    csv.AppendLine($"\"{blockName}\",{bm.BlockCount},{bm.PlacedCount},\"{note.Replace("\"", "\"\"")}\"");

//                    // Disable block lỗi
//                    bm.IsEnabled = false;
//                }

//                // Lưu file
//                try
//                {
//                    System.IO.File.WriteAllText(sfd.FileName, csv.ToString(), Encoding.UTF8);
//                    MessageBox.Show($"Đã xuất file lỗi thành công:\n{sfd.FileName}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Information);
//                }
//                catch (Exception ex)
//                {
//                    MessageBox.Show("Không thể lưu file CSV:\n" + ex.Message, "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
//                }
//            }
//        }

//        private void OkBtn_Click(object sender, RoutedEventArgs e)
//        {
//            if (DataContext is AutoPlaceElementFrBlockCADViewModel vm)
//                vm.RequestPlaceElements();
//        }

//        private void CancelBtn_Click(object sender, RoutedEventArgs e)
//        {
//            this.Close();
//        }

//        // Batch thao tác trên selected rows
//        private void BlockMappingGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
//        {
//            if (DataContext is AutoPlaceElementFrBlockCADViewModel vm)
//            {
//                vm.SelectedBlockMappings = new System.Collections.ObjectModel.ObservableCollection<AutoPlaceElementFrBlockCADViewModel.BlockMapping>(
//                    BlockMappingGrid.SelectedItems.Cast<AutoPlaceElementFrBlockCADViewModel.BlockMapping>());
//            }
//        }

//        private void EnableSelected_Click(object sender, RoutedEventArgs e)
//        {
//            foreach (var bm in BlockMappingGrid.SelectedItems.OfType<AutoPlaceElementFrBlockCADViewModel.BlockMapping>())
//                bm.IsEnabled = true;
//        }

//        private void DisableSelected_Click(object sender, RoutedEventArgs e)
//        {
//            foreach (var bm in BlockMappingGrid.SelectedItems.OfType<AutoPlaceElementFrBlockCADViewModel.BlockMapping>())
//                bm.IsEnabled = false;
//        }

//        private void SetOffsetSelected_Click(object sender, RoutedEventArgs e)
//        {
//            if (double.TryParse(BatchOffsetBox.Text, out double offset))
//            {
//                foreach (var bm in BlockMappingGrid.SelectedItems.OfType<AutoPlaceElementFrBlockCADViewModel.BlockMapping>())
//                    bm.Offset = offset;
//            }
//        }

//        private void SetCategorySelected_Click(object sender, RoutedEventArgs e)
//        {
//            if (BatchCategoryComboBox.SelectedItem is Autodesk.Revit.DB.Category cat)
//            {
//                foreach (var bm in BlockMappingGrid.SelectedItems.OfType<AutoPlaceElementFrBlockCADViewModel.BlockMapping>())
//                    bm.SelectedCategoryMapping = cat;
//            }
//        }

//        private void SetFamilySelected_Click(object sender, RoutedEventArgs e)
//        {
//            if (BatchFamilyComboBox.SelectedItem is Autodesk.Revit.DB.Family fam)
//            {
//                foreach (var bm in BlockMappingGrid.SelectedItems.OfType<AutoPlaceElementFrBlockCADViewModel.BlockMapping>())
//                    bm.SelectedFamilyMapping = fam;
//            }
//        }

//        private void SetTypeSelected_Click(object sender, RoutedEventArgs e)
//        {
//            if (BatchTypeComboBox.SelectedItem is Autodesk.Revit.DB.FamilySymbol type)
//            {
//                foreach (var bm in BlockMappingGrid.SelectedItems.OfType<AutoPlaceElementFrBlockCADViewModel.BlockMapping>())
//                    bm.SelectedTypeSymbolMapping = type;
//            }
//        }
//    }
//}
