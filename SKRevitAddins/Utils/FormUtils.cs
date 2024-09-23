using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace SKRevitAddins.Utils
{
    public class FormUtils
    {
        private static JtWindowHandle windowHandle = null;

        public static JtWindowHandle GetMainWindow()
        {
            if (windowHandle == null)
            {
                Process process = Process.GetCurrentProcess();
                IntPtr h = process.MainWindowHandle;
                windowHandle =  new JtWindowHandle(h);
            }
            return windowHandle;
        }
        public class JtWindowHandle : IWin32Window
        {
            IntPtr _hwnd;

            public JtWindowHandle(IntPtr h)
            {
                Debug.Assert(IntPtr.Zero != h,
                  "expected non-null window handle");

                _hwnd = h;
            }

            public IntPtr Handle
            {
                get
                {
                    return _hwnd;
                }
            }
        }
    }

}
