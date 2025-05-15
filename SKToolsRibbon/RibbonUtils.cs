// RibbonUtils.cs
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;  // cho BitmapImage

namespace SKToolsRibbon
{
    /// <summary>
    /// Hỗ trợ tạo PushButtonData với icon và tooltip.
    /// </summary>
    public class RibbonUtils
    {
        private readonly ControlledApplication _app;

        // 1) Thêm constructor như bạn gọi trong Ribbon.cs
        public RibbonUtils(ControlledApplication app)
        {
            _app = app;
        }

        /// <summary>
        /// Tạo PushButtonData nếu tìm thấy DLL và icon.
        /// </summary>
        public PushButtonData CreatePushButtonData(
            string name,
            string text,
            string assemblyPath,
            string className,
            string iconName,
            string tooltip = null)
        {
            // Nếu DLL không tồn tại thì không tạo button
            if (!File.Exists(assemblyPath))
                return null;

            // Khởi tạo PushButtonData
            var data = new PushButtonData(name, text, assemblyPath, className);

            // 2) Load icon nếu có
            if (!string.IsNullOrEmpty(iconName))
            {
                // Lấy folder chứa DLL
                var dllDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                // Lui lên 1 cấp để ra SKTools.bundle
                var bundleDir = Path.GetDirectoryName(dllDir);
                // Đường dẫn tới thư mục Icon
                var iconPath = Path.Combine(bundleDir, "Icon", iconName);

                if (File.Exists(iconPath))
                {
                    // BitmapImage với Uri (System.Uri, System.UriKind)
                    var img = new BitmapImage(new Uri(iconPath, UriKind.Absolute));
                    data.LargeImage = img;
                    data.Image = img;
                }
            }

            // Tooltip nếu có
            if (!string.IsNullOrEmpty(tooltip))
                data.ToolTip = tooltip;

            return data;
        }
    }
}
