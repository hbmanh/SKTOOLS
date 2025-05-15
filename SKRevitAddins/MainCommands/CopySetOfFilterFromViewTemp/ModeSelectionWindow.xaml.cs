using System.Windows;
using Autodesk.Revit.UI;
using SKRevitAddins.ViewModel;
using SKRevitAddins.Commands.CopySetOfFilterFromViewTemp;

namespace SKRevitAddins.Forms
{
    public partial class ModeSelectionWindow : Window
    {
        private UIApplication _uiApp;

        public ModeSelectionWindow(UIApplication uiApp)
        {
            InitializeComponent();
            _uiApp = uiApp;
        }

        private void TemplateToViews_Click(object sender, RoutedEventArgs e)
        {
            var viewModel = new CopySetOfFilterFromViewTempViewModel(_uiApp);
            App.thisApp.ShowCopySetFilterFromViewTempViewModel(_uiApp, viewModel);
            this.Close();
        }

        private void TemplateToTemplate_Click(object sender, RoutedEventArgs e)
        {
            var window = new CopySetOfFilterFromTemplateToTemplateWpfWindow(_uiApp);
            window.ShowDialog();
            this.Close();
        }
    }
}