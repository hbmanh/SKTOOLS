using System;
using System.IO;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace SKRevitAddins.Utils
{
    public static class LogoHelper
    {
        public static string GetLogoPath()
        {
            try
            {
                string asm = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string asmDir = Path.GetDirectoryName(asm);

                string p1 = Path.Combine(asmDir, "Icon", "logo.png");
                if (File.Exists(p1)) return p1;

                string dir = asmDir;
                for (int i = 0; i < 4 && !string.IsNullOrEmpty(dir); i++)
                {
                    try
                    {
                        var bundles = Directory.GetDirectories(dir, "*.bundle", SearchOption.TopDirectoryOnly);
                        foreach (var b in bundles)
                        {
                            string p = Path.Combine(b, "Icon", "logo.png");
                            if (File.Exists(p)) return p;
                        }
                    }
                    catch { }

                    dir = Path.GetDirectoryName(dir);
                }
            }
            catch { }

            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string addinsRoot = Path.Combine(appData, "Autodesk", "Revit", "Addins");
                if (Directory.Exists(addinsRoot))
                {
                    foreach (var year in Directory.GetDirectories(addinsRoot))
                    {
                        string p = Path.Combine(year, "SKTools.bundle", "Icon", "logo.png");
                        if (File.Exists(p)) return p;

                        var bundles = Directory.GetDirectories(year, "SKTools.bundle", SearchOption.TopDirectoryOnly);
                        foreach (var b in bundles)
                        {
                            string pb = Path.Combine(b, "Icon", "logo.png");
                            if (File.Exists(pb)) return pb;
                        }
                    }
                }
            }
            catch { }

            try
            {
                string fallback = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "SKRevitAddins",
                    "logo.png");

                if (File.Exists(fallback))
                    return fallback;
            }
            catch { }

            return null;
        }

        public static void TryLoadLogo(Image img)
        {
            try
            {
                string path = GetLogoPath();
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    var bi = new BitmapImage();
                    bi.BeginInit();
                    bi.UriSource = new Uri(path, UriKind.Absolute);
                    bi.CacheOption = BitmapCacheOption.OnLoad;
                    bi.EndInit();
                    img.Source = bi;
                }
            }
            catch
            {
            }
        }
    }
}
