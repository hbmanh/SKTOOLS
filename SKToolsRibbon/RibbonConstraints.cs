using Autodesk.Revit.ApplicationServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKToolsRibbon
{
    public class RibbonConstraints
    {
        public string ContentsFolder;
        public string ResourcesFolder;
        public string HelpFolder;
        public string ImageFolder;
        public string DllFolder;

        public RibbonConstraints(ControlledApplication app = null)
        {
            ContentsFolder = @"C:\ProgramData\Autodesk\Revit\Addins\2022\SKTools.bundle\Contents";
            ResourcesFolder = Path.Combine(ContentsFolder, "Resources");
            HelpFolder = Path.Combine(ResourcesFolder, "Help");
            ImageFolder = Path.Combine(ResourcesFolder, "Images");

            if (app != null)
            {
                switch (app.VersionNumber)
                {
                    case "2017":
                        DllFolder = Path.Combine(ContentsFolder, "2017", "dll");
                        break;
                    case "2018":
                        DllFolder = Path.Combine(ContentsFolder, "2018", "dll");
                        break;
                    case "2019":
                        DllFolder = Path.Combine(ContentsFolder, "2019", "dll");
                        break;
                    case "2020":
                        DllFolder = Path.Combine(ContentsFolder, "2020", "dll");
                        break;
                    case "2021":
                        DllFolder = Path.Combine(ContentsFolder, "2021", "dll");
                        break;
                    case "2022":
                        DllFolder = Path.Combine(ContentsFolder, "2022", "dll");
                        break;
                    case "2023":
                        DllFolder = Path.Combine(ContentsFolder, "2023", "dll");
                        break;
                    case "2024":
                        DllFolder = Path.Combine(ContentsFolder, "2024", "dll");
                        break;
                    case "2025":
                        DllFolder = Path.Combine(ContentsFolder, "2025", "dll");
                        break;
                }
            }
        }
    }
}
