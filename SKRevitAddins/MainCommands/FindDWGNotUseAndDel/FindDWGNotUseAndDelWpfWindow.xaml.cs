using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;
using SKRevitAddins.Utils;

namespace SKRevitAddins.FindDWGNotUseAndDel
{
    public partial class FindDWGNotUseAndDelWpfWindow : Window
    {
        private readonly ExternalEvent _exEvent;
        private readonly FindDWGNotUseAndDelRequestHandler _handler;
        private readonly FindDWGNotUseAndDelViewModel _vm;

        public FindDWGNotUseAndDelWpfWindow(
            ExternalEvent exEvent,
            FindDWGNotUseAndDelRequestHandler handler,
            FindDWGNotUseAndDelViewModel viewModel)
        {
            InitializeComponent();

            _exEvent = exEvent;
            _handler = handler;
            _vm = viewModel;
            DataContext = _vm;

            LogoHelper.TryLoadLogo(LogoImage);
        }

        private void viewSetDg_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedDwgItems = viewSetDg.SelectedItems
                .OfType<FindDWGNotUseAndDelViewModel.DwgItem>()
                .ToList();

            _vm.SelectedDWGs = selectedDwgItems;
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string keyword = SearchTextBox.Text?.Trim();
            _vm.FilterDWGs(keyword);
        }

        private void DeleteBtn_Click(object sender, RoutedEventArgs e)
        {
            _vm.IsDeleteAll = AllRadioBtn.IsChecked == true;
            _handler.Request.Make(RequestId.Delete);
            _exEvent.Raise();
        }

        private void OpenViewBtn_Click(object sender, RoutedEventArgs e)
        {
            _handler.Request.Make(RequestId.OpenView);
            _exEvent.Raise();
        }

        private void ExportTableBtn_Click(object sender, RoutedEventArgs e)
        {
            _handler.Request.Make(RequestId.Export);
            _exEvent.Raise();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
