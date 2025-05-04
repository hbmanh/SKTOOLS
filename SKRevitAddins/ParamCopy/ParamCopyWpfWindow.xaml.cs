//using System.Windows;
//using Autodesk.Revit.UI;

//namespace ParamCopy
//{
//    /// <summary>
//    /// Interaction logic for ParamCopyWpfWindow.xaml
//    /// </summary>
//    public partial class ParamCopyWpfWindow : Window
//    {
//        private ParamCopyRequestHandler m_Handler;
//        private ExternalEvent m_ExEvent;
//        public ParamCopyWpfWindow(ExternalEvent exEvent,
//            ParamCopyRequestHandler handler,
//            ParamCopyViewModel paramCopyViewModel)
//        {
//            // string appPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
//            // var currentDirectory = System.IO.Path.GetDirectoryName(Assembly.GetCallingAssembly().Location);
//            // System.Reflection.Assembly.LoadFile(currentDirectory + "/Microsoft.Xaml.Behaviors.dll");

//            InitializeComponent();

//            // var _ = new Microsoft.Xaml.Behaviors.DefaultTriggerAttribute(typeof(Trigger), typeof(Microsoft.Xaml.Behaviors.TriggerBase), null);

//            m_Handler = handler;
//            m_ExEvent = exEvent;

//            this.LoadViewFromUri("/ParamCopy;component/ParamCopyWpfWindow.xaml");

//            this.DataContext = paramCopyViewModel;
//            // unitNoteLbl.Visibility = Visibility.Hidden;

//            instanceApplyBtn.Click += InstanceApplyBtn_Click;
//            familyApplyBtn.Click += FamilyApplyBtn_Click;
//            categoryApplyBtn.Click += CategoryApplyBtn_Click;
//            allEleApplyBtn.Click += AllEleApplyBtn_Click;

//            cancelBtn.Click += CancelBtn_Click;
//        }

//        private void AllEleApplyBtn_Click(object sender, RoutedEventArgs e)
//        {
//            MakeRequest(RequestId.AllEleCopy);
//        }

//        private void CategoryApplyBtn_Click(object sender, RoutedEventArgs e)
//        {
//            MakeRequest(RequestId.CategoryCopy);
//        }

//        private void FamilyApplyBtn_Click(object sender, RoutedEventArgs e)
//        {
//            MakeRequest(RequestId.FamilyCopy);
//        }

//        private void InstanceApplyBtn_Click(object sender, RoutedEventArgs e)
//        {
//            MakeRequest(RequestId.InstanceCopy);
//        }

//        private void CancelBtn_Click(object sender, RoutedEventArgs e)
//        {
//            this.Close();
//        }
//        private void MakeRequest(RequestId request)
//        {
//            m_Handler.Request.Make(request);
//            m_ExEvent.Raise();
//        }
//    }
//}
