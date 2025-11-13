using System;
using System.IO;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Forms; // Thêm namespace cho WinForms
using PictureBox = System.Windows.Forms.PictureBox; // Aliasing để dùng PictureBox trong WinForms

namespace SKRevitAddins.Utils
{
    public static class LogoHelper
    {
        public static string GetLogoPath()
        {
            try
            {
                string asmPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string asmDir = Path.GetDirectoryName(asmPath);

                string p1 = Path.Combine(asmDir, "Icon", "logo.png");
                if (File.Exists(p1)) return p1;

                string dir = asmDir;
                for (int i = 0; i < 4 && !string.IsNullOrEmpty(dir); i++)
                {
                    try
                    {
                        var bundles = Directory.GetDirectories(dir, "*.bundle", SearchOption.TopDirectoryOnly);
                        foreach (var bundle in bundles)
                        {
                            string p = Path.Combine(bundle, "Icon", "logo.png");
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
                    foreach (var yearDir in Directory.GetDirectories(addinsRoot))
                    {
                        string candidate = Path.Combine(yearDir, "SKTools.bundle", "Icon", "logo.png");
                        if (File.Exists(candidate)) return candidate;

                        var bundles = Directory.GetDirectories(yearDir, "SKTools.bundle", SearchOption.TopDirectoryOnly);
                        foreach (var bundle in bundles)
                        {
                            string p = Path.Combine(bundle, "Icon", "logo.png");
                            if (File.Exists(p)) return p;
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

                if (File.Exists(fallback)) return fallback;
            }
            catch { }

            return null;
        }

        public static void TryLoadLogo(object img)
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

                    if (img is Image wpfImage)
                    {
                        wpfImage.Source = bi; // WPF Image
                    }
                    else if (img is PictureBox winFormsPictureBox)
                    {
                        winFormsPictureBox.Image = System.Drawing.Image.FromFile(path); // WinForms PictureBox
                    }
                }
            }
            catch
            {
            }
        }
    }
}
