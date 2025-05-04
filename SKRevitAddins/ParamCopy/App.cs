//using Autodesk.Revit.UI;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Interop;
//using System.Windows.Media.Imaging;
//using System.Windows;
//using System.IO;
//using D3Lib;

//namespace ParamCopy
//{
//    public class App : IExternalApplication
//    {
//        internal static App thisApp = null;
//        private ParamCopyWpfWindow m_ParamCopyWpfWindow;
//        private PushButton m_paramCopyBtn;
//        public Result OnShutdown(UIControlledApplication application)
//        {
//            if (m_ParamCopyWpfWindow != null && m_ParamCopyWpfWindow.IsVisible)
//            {
//                m_ParamCopyWpfWindow.Close();
//            }
//            return Result.Succeeded;
//        }

//        public Result OnStartup(UIControlledApplication application)
//        {
//            m_ParamCopyWpfWindow = null;
//            thisApp = this;

//            var versionText = application.ControlledApplication.VersionNumber;
//            var version = int.Parse(versionText);
//            var theme = UIThemeManager.CurrentTheme;

//            string tabName = "検証(KD)";
//            string panelName = "***";

//            switch (CheckLib.CurrentMode())
//            {
//                case Mode.INVALID:
//                    break;
//                case Mode.KAD:
//                    tabName = "建築(KAD)";
//                    panelName = "AAA";
//                    break;
//                case Mode.KSD:
//                    tabName = "***(KSD)";
//                    panelName = "SSS";
//                    break;
//                case Mode.KMD:
//                    tabName = "設備(KMD)";
//                    panelName = "MMM";
//                    break;
//            }

//            try
//            {
//                application.CreateRibbonTab(tabName);
//            }
//            catch
//            {
//            }

//            var panelList = application.GetRibbonPanels(tabName);

//            int index = panelList.FindIndex(r => r.Name == panelName);

//            RibbonPanel panel = index == -1 ? application.CreateRibbonPanel(tabName, panelName) : panelList[index];


//            System.Drawing.Bitmap paramCopyIcon32 = null;
//            System.Drawing.Bitmap paramCopyIcon16 = null;

//            if (version < 2024)
//            {
//                paramCopyIcon32 = ParamCopy.Properties.Resources.コピー_パラメータ_32px_R22;
//                paramCopyIcon16 = ParamCopy.Properties.Resources.コピー_パラメータ_16px_R22;
//            }
//            else if (theme == UITheme.Dark)
//            {
//                paramCopyIcon32 = ParamCopy.Properties.Resources.コピー_パラメータ_32px_R24_暗;
//                paramCopyIcon16 = ParamCopy.Properties.Resources.コピー_パラメータ_16px_R24_暗;
//            }
//            else
//            {
//                paramCopyIcon32 = ParamCopy.Properties.Resources.コピー_パラメータ_32px_R24_明;
//                paramCopyIcon16 = ParamCopy.Properties.Resources.コピー_パラメータ_16px_R24_明;
//            }

//            PushButtonData paramCopyBtn
//                = CreatePushButtonDataFromBitmap("パラメータ間コピー",
//                "パラメータ間\nコピー", "ParamCopy.dll",
//                "ParamCopy.ParamCopyCmd", paramCopyIcon32,
//                "パラメータコピー",""
//                , null, null, null, paramCopyIcon16);
//            //panel.AddItem(paramCopyBtn);
//            m_paramCopyBtn = panel.AddItem(paramCopyBtn) as PushButton;
//#if REVIT2024
//            application.ThemeChanged += Application_ThemeChanged;
//#endif
//            return Result.Succeeded;
//        }
//#if REVIT2024
//        private void Application_ThemeChanged(object sender, Autodesk.Revit.UI.Events.ThemeChangedEventArgs e)
//        {
//            UpdateImageByTheme();
//        }
//        private void UpdateImageByTheme()
//        {
//            System.Drawing.Bitmap paramCopyIcon32 = null;
//            System.Drawing.Bitmap paramCopyIcon16 = null;

//            UITheme theme = UIThemeManager.CurrentTheme;
//            switch (theme)
//            {
//                case UITheme.Dark:
//                    paramCopyIcon32 = ParamCopy.Properties.Resources.コピー_パラメータ_32px_R24_暗;
//                    paramCopyIcon16 = ParamCopy.Properties.Resources.コピー_パラメータ_16px_R24_暗;
//                    break;
//                case UITheme.Light:
//                    paramCopyIcon32 = ParamCopy.Properties.Resources.コピー_パラメータ_32px_R24_明;
//                    paramCopyIcon16 = ParamCopy.Properties.Resources.コピー_パラメータ_16px_R24_明;
//                    break;
//            }
//            m_paramCopyBtn.Image = Imaging.CreateBitmapSourceFromHBitmap(paramCopyIcon16.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
//            m_paramCopyBtn.LargeImage = Imaging.CreateBitmapSourceFromHBitmap(paramCopyIcon32.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
//        }
//#endif
//        public void ShowParamCopyWindow(UIApplication uiapp, ParamCopyViewModel viewModel)
//        {
//            if (m_ParamCopyWpfWindow == null || !m_ParamCopyWpfWindow.IsVisible)
//            {
//                ParamCopyRequestHandler handler = new ParamCopyRequestHandler(viewModel);
//                ExternalEvent exEvent = ExternalEvent.Create(handler);
//                m_ParamCopyWpfWindow = new ParamCopyWpfWindow(exEvent, handler, viewModel);
//                m_ParamCopyWpfWindow.Show();
//            }
//        }

//        public PushButtonData CreatePushButtonDataFromBitmap(
//           string name, string displayName,
//           string dllName, string fullClassName,
//           System.Drawing.Bitmap largeImage, string tooltip,
//           string helperPath = null,
//           string longDescription = null,
//           string tooltipImage = null,
//           string linkYoutube = null,
//           System.Drawing.Bitmap image = null)
//        {
//            string directoryName = Path.GetDirectoryName(this.GetType().Assembly.Location);
//            try
//            {
//                PushButtonData pushButtonData = new PushButtonData(
//                    name, displayName,
//                    Path.Combine(directoryName, dllName),
//                    fullClassName);
//                var largeImageHandle = largeImage.GetHbitmap();
//                var convertedLargeImage = Imaging.CreateBitmapSourceFromHBitmap(largeImageHandle, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
//                pushButtonData.LargeImage = convertedLargeImage;

//                if (image != null)
//                {
//                    var imageHandle = image.GetHbitmap();
//                    var convertedImage = Imaging.CreateBitmapSourceFromHBitmap(imageHandle, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
//                    pushButtonData.Image = convertedImage;
//                }

//                pushButtonData.ToolTip = tooltip;

//                //if (!string.IsNullOrEmpty(tooltipImage))
//                //{
//                //    Uri tooltipUri = new Uri(Path.Combine(imageFolder, tooltipImage),
//                //        UriKind.Absolute);
//                //    pushButtonData.ToolTipImage = new BitmapImage(tooltipUri);
//                //}

//                if (!string.IsNullOrEmpty(helperPath))
//                {
//                    ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.ChmFile,
//                        helperPath);
//                    pushButtonData.SetContextualHelp(contextHelp);
//                }

//                if (!string.IsNullOrEmpty(linkYoutube))
//                {
//                    ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.Url,
//                        linkYoutube);
//                    pushButtonData.SetContextualHelp(contextHelp);
//                }

//                if (longDescription != null)
//                {
//                    pushButtonData.LongDescription = longDescription;
//                }
//                return pushButtonData;
//            }
//            catch (Exception)
//            {
//                // TaskDialog.Show("Error Project", name + ", " + displayName);
//                MessageBox.Show(name + ", " + displayName, "エラー");
//            }
//            return null;
//        }
//    }
//}
