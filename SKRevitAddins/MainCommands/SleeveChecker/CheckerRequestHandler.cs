using Autodesk.Revit.UI;

namespace SKRevitAddins.SleeveChecker
{
    public class CheckerRequestHandler : IExternalEventHandler
    {
        private readonly SleeveCheckerViewModel _viewModel;

        public CheckerRequestHandler(SleeveCheckerViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        public void Execute(UIApplication app)
        {
            switch (_viewModel.PendingRequest)
            {
                case SleeveCheckerViewModel.RequestId.Preview:
                    _viewModel.DoPreview(app);
                    break;
                case SleeveCheckerViewModel.RequestId.Apply:
                    _viewModel.DoApply(app);
                    break;
                case SleeveCheckerViewModel.RequestId.ShowError:
                    _viewModel.DoShowError(app);
                    break;
                default:
                    break;
            }
            _viewModel.PendingRequest = SleeveCheckerViewModel.RequestId.None;
        }

        public string GetName() => "SleeveChecker ExternalEventHandler";
    }
}