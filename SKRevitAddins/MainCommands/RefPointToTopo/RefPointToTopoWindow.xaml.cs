using System.Windows;
using Autodesk.Revit.UI;

namespace SKRevitAddins.RefPointToTopo
{
    public partial class RefPointToTopoWindow : Window
    {
        private readonly RefPointToTopoViewModel _vm;

        public RefPointToTopoWindow(UIDocument uiDoc)
        {
            InitializeComponent();
            _vm = new RefPointToTopoViewModel(uiDoc);
            DataContext = _vm;

            Loaded += (_, __) => this.Focus();
        }

        protected override void OnPreviewKeyDown(System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
            {
                e.Handled = true;
                Close();
            }
            base.OnPreviewKeyDown(e);
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (_vm?.IsBusy == true)
            {
                var r = MessageBox.Show(
                    "Đang xử lý. Bạn có muốn hủy và đóng cửa sổ không?",
                    "Xác nhận",
                    MessageBoxButton.YesNo, MessageBoxImage.Question);

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
