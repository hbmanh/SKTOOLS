using System.Windows;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    public partial class SheetSelectionWindow : Window
    {
        public SheetSelectionWindow(LayoutsToDWGViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;
        }

        private void OK_Click(object sender, RoutedEventArgs e) => Close();
        private void Cancel_Click(object sender, RoutedEventArgs e) => Close();
    }
}
