using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;

namespace SKRevitAddins.PointCloudAddins.RefPointToTopo
{
    public partial class RefPointToTopoWindow : Window
    {
        readonly RefPointToTopoViewModel _vm;

        public RefPointToTopoWindow(UIDocument uiDoc)
        {
            InitializeComponent();
            _vm = new RefPointToTopoViewModel(uiDoc);
            DataContext = _vm;

            Loaded += (_, __) =>
            {
                LogoHelper.TryLoadLogo(LogoImage);
                Focus();
            };
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                Close();
            }
            base.OnPreviewKeyDown(e);
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (_vm?.IsBusy == true)
            {
                var r = MessageBox.Show(
                    "Đang xử lý. Bạn có muốn hủy và đóng cửa sổ không?",
                    "Xác nhận",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (r == MessageBoxResult.No)
                {
                    e.Cancel = true;
                    return;
                }

                _vm.CancelCmd?.Execute(null);
            }

            base.OnClosing(e);
        }
    }
}
