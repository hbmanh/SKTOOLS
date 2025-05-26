using Autodesk.Revit.UI;
using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media.Imaging;

namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
{
    public partial class PlaceElementsFromBlocksCadWpfWindow : Window
    {
        public PlaceElementsFromBlocksCadWpfWindow(
            ExternalEvent exEvent,
            AutoPlaceElementFrBlockCADRequestHandler handler,
            AutoPlaceElementFrBlockCADViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;

            viewModel.UpdateStatus = msg =>
            {
                Dispatcher.Invoke(() => StatusText.Text = msg);
            };

            try
            {
                // ✅ Sửa lỗi CS0117: Lấy version từ viewModel.UiApp
                string version = viewModel.UiApp.Application.VersionNumber;
                string iconPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Autodesk", "Revit", "Addins", version, "SKTools.bundle", "Icon", "shinken.png");

                if (System.IO.File.Exists(iconPath))
                {
                    LogoImage.Source = new BitmapImage(new Uri(iconPath, UriKind.Absolute));
                }
            }
            catch
            {
                // Bạn có thể log lỗi ở đây nếu cần
            }
        }

        private void EnableAll_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is AutoPlaceElementFrBlockCADViewModel vm)
                foreach (var bm in vm.BlockMappings)
                    bm.IsEnabled = true;
        }

        private void DisableAll_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is AutoPlaceElementFrBlockCADViewModel vm)
                foreach (var bm in vm.BlockMappings)
                    bm.IsEnabled = false;
        }

        private void ShowErrors_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is AutoPlaceElementFrBlockCADViewModel vm)
            {
                var failed = vm.BlockMappings
                    .Where(bm => bm.HasPlacementRun && bm.PlacedCount != bm.BlockCount)
                    .ToList();

                if (failed.Count == 0)
                {
                    MessageBox.Show("Không có block nào bị lỗi!", "Note", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var sb = new StringBuilder();
                sb.AppendLine("Các Block bị lỗi khi đặt Family Instance:");
                sb.AppendLine();

                foreach (var bm in failed)
                {
                    string blockName = bm.DisplayBlockName;
                    string note = string.IsNullOrWhiteSpace(bm.FailureNote)
                        ? $"Chỉ đặt được {bm.PlacedCount}/{bm.BlockCount} instance."
                        : bm.FailureNote;

                    sb.AppendLine($"{blockName}: {note}");
                }

                MessageBox.Show(sb.ToString(), "Block Error Notes", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is AutoPlaceElementFrBlockCADViewModel vm)
                vm.RequestPlaceElements();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
