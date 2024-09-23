using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Windows.Media.Imaging;

using AW = Autodesk.Windows;

namespace SKToolsRibbon
{
    public class RibbonUtils
    {
        private RibbonConstraints ribbonConstraints;
        private string imageFolder;
        private string dllFolder;

        public RibbonUtils(ControlledApplication app)
        {
            ribbonConstraints = new RibbonConstraints(app);
            imageFolder = ribbonConstraints.ImageFolder;
            dllFolder = ribbonConstraints.DllFolder;
        }

        public PushButtonData CreatePushButtonData(
            string name, string displayName,
            string dllName, string fullClassName,
            string largeImage, string tooltip,
            string helperPath = null,
            string longDescription = null,
            string tooltipImage = null,
            string linkYoutube = null,
            string image = null)
        {
            try
            {
                PushButtonData pushButtonData = new PushButtonData(
                    name, displayName,
                    Path.Combine(dllFolder, dllName),
                    fullClassName);

                pushButtonData.LargeImage = CreateBitmapImage(imageFolder, largeImage);

                if (image != null)
                {
                    pushButtonData.Image = CreateBitmapImage(imageFolder, image);
                }

                pushButtonData.ToolTip = tooltip;

                if (!string.IsNullOrEmpty(tooltipImage))
                {
                    Uri tooltipUri = new Uri(Path.Combine(imageFolder, tooltipImage),
                        UriKind.Absolute);
                    pushButtonData.ToolTipImage = new BitmapImage(tooltipUri);
                }

                if (!string.IsNullOrEmpty(helperPath))
                {
                    ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.ChmFile,
                        helperPath);
                    pushButtonData.SetContextualHelp(contextHelp);
                }

                if (!string.IsNullOrEmpty(linkYoutube))
                {
                    ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.Url,
                        linkYoutube);
                    pushButtonData.SetContextualHelp(contextHelp);
                }

                if (longDescription != null)
                {
                    pushButtonData.LongDescription = longDescription;
                }

                return pushButtonData;
            }
            catch (Exception)
            {
                TaskDialog.Show("Error Project", name + ", " + displayName);
            }
            return null;
        }

        public PulldownButtonData CreatePulldownButtonData(
            string name,
            string displayName,
            string tooltip,
            string image,
            string helperPath = null,
            string linkYoutube = null,
            string toolTipImage = null)
        {

            PulldownButtonData pulldownButtonData = pulldownButtonData
                = new PulldownButtonData(name, displayName);

            pulldownButtonData.ToolTip = tooltip;

            if (!string.IsNullOrEmpty(helperPath))
            {
                ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.ChmFile,
                    helperPath);
                pulldownButtonData.SetContextualHelp(contextHelp);
            }
            if (!string.IsNullOrEmpty(linkYoutube))
            {
                ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.Url,
                    linkYoutube);
                pulldownButtonData.SetContextualHelp(contextHelp);
            }

            if (!string.IsNullOrEmpty(image))
            {
                pulldownButtonData.Image
                    = CreateBitmapImage(imageFolder, image);
            }

            if (!string.IsNullOrEmpty(toolTipImage))
            {
                pulldownButtonData.ToolTipImage
                    = CreateBitmapImage(imageFolder, toolTipImage);
            }

            return pulldownButtonData;
        }

        public PulldownButton CreatePulldownButton(
            RibbonPanel ribbonPanel,
            string name, string displayName,
            string tooltip,
            string largeImage,
            string toolTipImage = null,
            string helperPath = null,
            string linkYoutube = null)
        {

            PulldownButtonData pulldownButtonData
                = new PulldownButtonData(name, displayName);
            pulldownButtonData.ToolTip = tooltip;


            if (!string.IsNullOrEmpty(helperPath))
            {
                ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.ChmFile,
                    helperPath);
                pulldownButtonData.SetContextualHelp(contextHelp);
            }
            if (!string.IsNullOrEmpty(linkYoutube))
            {
                ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.Url,
                    linkYoutube);
                pulldownButtonData.SetContextualHelp(contextHelp);
            }

            if (!string.IsNullOrEmpty(largeImage))
            {
                pulldownButtonData.LargeImage
                    = CreateBitmapImage(imageFolder, largeImage);
            }

            if (!string.IsNullOrEmpty(toolTipImage))
            {
                pulldownButtonData.ToolTipImage
                    = CreateBitmapImage(imageFolder, toolTipImage);
            }

            return ribbonPanel.AddItem(pulldownButtonData) as PulldownButton;
        }

        public SplitButton CreateSplitButton(
            RibbonPanel ribbonPanel,
            string name,
            string displayName,
            string tooltip,
            string helperPath = null,
            string linkYoutube = null)
        {
            SplitButtonData splitButtonData = new SplitButtonData(name, displayName);
            splitButtonData.ToolTip = tooltip;

            if (!string.IsNullOrEmpty(helperPath))
            {
                ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.ChmFile,
                        helperPath);
                splitButtonData.SetContextualHelp(contextHelp);
            }
            if (!string.IsNullOrEmpty(linkYoutube))
            {
                ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.Url,
                    linkYoutube);
                splitButtonData.SetContextualHelp(contextHelp);
            }

            return ribbonPanel.AddItem(splitButtonData) as SplitButton;
        }

        public BitmapImage CreateBitmapImage(string imageFolder, string image)
        {
            string pathImage = Path.Combine(imageFolder, image);
            Uri iconUri = new Uri(pathImage, UriKind.Absolute);
            return new BitmapImage(iconUri);
        }


    }
    public class RibbonUtils2
    {
        public AW.RibbonItem GetButton(string tabName, string panelName, string itemName)
        {
            AW.RibbonControl ribbon = AW.ComponentManager.Ribbon;
            foreach (AW.RibbonTab tab in ribbon.Tabs)
            {
                if (tab.Name == tabName)
                {
                    foreach (AW.RibbonPanel panel in tab.Panels)
                    {
                        if (panel.Source.Title == panelName)
                        {
                            return panel.FindItem("CustomCtrl_%CustomCtrl_%" + tabName + "%" + panelName + "%" + itemName, true) as AW.RibbonItem;
                        }
                    }
                }
            }
            return null;
        }
    }
}
