using Autodesk.Revit.UI;
using System.Windows;
using System.Windows.Input;
using static SKRevitAddins.AutoPlaceElementFrBlockCAD.AutoPlaceElementFrBlockCADViewModel;

namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
{
    public partial class PlaceElementsFromBlocksCadWpfWindow : Window
    {
        private AutoPlaceElementFrBlockCADRequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        private int lastCheckedIndex = -1;

        public PlaceElementsFromBlocksCadWpfWindow(
            ExternalEvent exEvent,
            AutoPlaceElementFrBlockCADRequestHandler handler,
            AutoPlaceElementFrBlockCADViewModel viewModel)
        {
            InitializeComponent();
            m_Handler = handler;
            m_ExEvent = exEvent;
            DataContext = viewModel;

            // Cập nhật trạng thái tiến trình
            viewModel.UpdateStatus = (msg) =>
            {
                Dispatcher.Invoke(() => StatusText.Text = msg);
            };
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

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            m_Handler.Request.Make(RequestId.OK);
            m_ExEvent.Raise();
            this.Close();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void CheckBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
            {
                var vm = DataContext as AutoPlaceElementFrBlockCADViewModel;
                if (vm == null) return;

                var item = (sender as FrameworkElement)?.DataContext as BlockMapping;
                int currIndex = vm.BlockMappings.IndexOf(item);
                if (lastCheckedIndex == -1)
                {
                    lastCheckedIndex = currIndex;
                    return;
                }
                int from = System.Math.Min(lastCheckedIndex, currIndex);
                int to = System.Math.Max(lastCheckedIndex, currIndex);
                bool? toSet = !(item.IsEnabled);
                for (int i = from; i <= to; i++)
                    vm.BlockMappings[i].IsEnabled = toSet.Value;

                lastCheckedIndex = currIndex;
                e.Handled = true;
            }
            else
            {
                var vm = DataContext as AutoPlaceElementFrBlockCADViewModel;
                var item = (sender as FrameworkElement)?.DataContext as BlockMapping;
                if (vm != null && item != null)
                    lastCheckedIndex = vm.BlockMappings.IndexOf(item);
            }
        }
    }
}
