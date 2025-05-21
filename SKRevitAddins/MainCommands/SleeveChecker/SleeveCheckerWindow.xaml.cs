using Autodesk.Revit.UI;
using System.Windows;

namespace SKRevitAddins.SleeveChecker
{
    public partial class SleeveCheckerWindow : Window
    {
        public SleeveCheckerWindow(ExternalEvent exEvent, CheckerRequestHandler handler, SleeveCheckerViewModel viewModel)
        {
            InitializeComponent();
            viewModel.SetExternalEvent(exEvent);
            DataContext = viewModel;
        }
    }
}