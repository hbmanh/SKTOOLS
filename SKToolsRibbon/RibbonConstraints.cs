using System;
using System.IO;

namespace SKToolsRibbon
{
    public class RibbonConstraints
    {
        public string BundleFolder { get; }
        public string DllFolder { get; }
        public string HelpFolder { get; }
        public string IconFolder { get; }

        public RibbonConstraints()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            BundleFolder = Path.Combine(baseDir, "SKTools.bundle");
            DllFolder = Path.Combine(BundleFolder, "dll");
            HelpFolder = Path.Combine(BundleFolder, "help");
            IconFolder = Path.Combine(BundleFolder, "Icon");
        }
    }
}