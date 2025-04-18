using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    public class RunAutoCadBg
    {
        private readonly bool _openFile;
        private readonly bool _silent;
        private readonly string _scriptFile;
        private readonly string _outputFile;

        /// <summary>
        /// openFile: sau khi merge xong có mở thư mục không
        /// scriptFile: đường dẫn tới .scr do chúng ta sinh ra (AcadAppFolder.AcadCmdFile)
        /// outputFile: đường dẫn file DWG cuối cùng
        /// silent: true chạy ẩn, false bật Visible
        /// </summary>
        public RunAutoCadBg(bool openFile, string scriptFile, string outputFile, bool silent = true)
        {
            _openFile = openFile;
            _silent = silent;
            _scriptFile = scriptFile;
            _outputFile = outputFile;

            try
            {
                var acad = LaunchAutoCAD();
                if (acad == null)
                {
                    MessageBox.Show(
                        "Không tìm thấy AutoCAD qua COM. Hãy chắc bạn đã cài AutoCAD và chạy với quyền đủ.",
                        "Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                acad.Visible = !_silent;

                // gửi lệnh chạy script
                // .SCRIPT thì phải bắt đầu với dấu gạch ngang nếu trong tiếng Việt nó auto‑translate
                acad.ActiveDocument.SendCommand($"-SCRIPT\n\"{_scriptFile}\"\n");

                if (_silent)
                {
                    // chờ file DWG đầu ra xuất hiện
                    WaitForFile(_outputFile, TimeSpan.FromSeconds(60));
                    if (_openFile)
                        Process.Start("explorer.exe", Path.GetDirectoryName(_outputFile));
                    acad.Quit();
                }
                else
                {
                    // nếu không silent, đợi user tự tương tác hoặc bắt sự kiện Closed (khó với COM)
                    // ở đây ta chỉ schedule mở folder khi script xong:
                    Task.Run(async () =>
                    {
                        await Task.Delay(30_000); // mong là script kịp chạy xong
                        if (_openFile && File.Exists(_outputFile))
                            Process.Start("explorer.exe", Path.GetDirectoryName(_outputFile));
                    });
                }
            }
            catch (COMException comEx)
            {
                MessageBox.Show(
                    "COMException: " + comEx.Message,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Exception: " + ex.Message,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private dynamic LaunchAutoCAD()
        {
            // Danh sách ProgID AutoCAD (thử cao xuống thấp)
            string[] progIds = {
                "AutoCAD.Application.28", // AutoCAD 2024
                "AutoCAD.Application.24", // AutoCAD 2022
                "AutoCAD.Application.22", // AutoCAD 2020
                "AutoCAD.Application"     // default
            };

            foreach (var pid in progIds)
            {
                try
                {
                    Type acadType = Type.GetTypeFromProgID(pid);
                    if (acadType != null)
                    {
                        dynamic acad = Activator.CreateInstance(acadType);
                        return acad;
                    }
                }
                catch
                {
                    // bỏ qua, thử ProgID kế
                }
            }
            return null;
        }

        private void WaitForFile(string path, TimeSpan timeout)
        {
            var sw = Stopwatch.StartNew();
            while (sw.Elapsed < timeout)
            {
                if (File.Exists(path)) return;
                Thread.Sleep(500);
            }
        }
    }
}
