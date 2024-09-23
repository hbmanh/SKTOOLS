using System.Windows;
using SKRevitAddins.ViewModel;
using Window = System.Windows.Window;

namespace SKRevitAddins.Forms
{
    public partial class SelectElementsVer1ReplaceTextWpfWindow : Window
    {
        public SelectElementsVer1ReplaceTextWpfWindow(SelectElementsVer1ViewModel viewModel)
        {
            InitializeComponent();
            this.DataContext = viewModel;
        }
        
        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }
        private void Numbering_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


    }
}
