using Autodesk.Revit.DB;
using Autodesk.Revit.DB.PDF;
using Autodesk.Revit.UI;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Input;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace SKRevitAddins.LayoutsToDWG
{
    //──────────────────────── General utils ───────────────────────
    internal static class Util
    {
        /// <summary>Dictionary.GetOrAdd (tương tự .NET 6).</summary>
        public static TValue GetOrAdd<TKey, TValue>(this IDictionary<TKey, TValue> dic, TKey key)
            where TValue : new()
        {
            if (!dic.TryGetValue(key, out var v))
            {
                v = new TValue();
                dic[key] = v;
            }
            return v;
        }

        /// <summary>Loại bỏ ký tự cấm trong tên file.</summary>
        public static string Sanitize(string s)
        {
            foreach (char c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
            return s.Trim();
        }
    }

    //──────────────────────── Relay & base VM ─────────────────────
    public class RelayCommand : ICommand
    {
        readonly Action<object> _a; readonly Func<object, bool> _c;
        public RelayCommand(Action<object> a, Func<object, bool> c = null) { _a = a; _c = c; }
        public bool CanExecute(object p) => _c?.Invoke(p) ?? true;
        public void Execute(object p) => _a(p);
        public event EventHandler CanExecuteChanged;
        public void Raise() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }

    public abstract class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    //──────────────────────── PDF attach data ─────────────────────
    internal record PdfData(
        string FileName, int Page,
        double InsXmm, double InsYmm,
        double Scale, double RotationDeg);

    //──────────────────────── Layer helper ────────────────────────
    internal static class LayerExportHelper
    {
        public static void WriteLayerMapping(Document doc, string setup, string txtPath)
        {
            var opt = DWGExportOptions.GetPredefinedOptions(doc, setup)
                      ?? throw new InvalidOperationException($"Không tìm thấy setup “{setup}”.");
            var tbl = opt.GetExportLayerTable();
            if (tbl == null || !tbl.GetKeys().Any())
                throw new InvalidOperationException("Setup này không có layer‑mapping.");

            var sb = new StringBuilder(4096)
                     .AppendLine("# Category\tSubcategory\tLayer\tColor");
            foreach (var k in tbl.GetKeys())
            {
                var i = tbl[k];
                if (i.ColorNumber < 0) continue;
                sb.Append(k.CategoryName).Append('\t')
                  .Append(k.SubCategoryName ?? "").Append('\t')
                  .Append(i.LayerName).Append('\t')
                  .Append(i.ColorNumber).AppendLine();
            }
            File.WriteAllText(txtPath, sb.ToString(), Encoding.UTF8);
        }
    }

    //──────────────────────── Build & run AutoCAD script ──────────
    internal static class PdfAttachScriptBuilder
    {
        const double Ft2mm = 304.8;

        public static string BuildScript(string dir,
                                         IReadOnlyDictionary<string, List<PdfData>> map)
        {
            string scr = Path.Combine(dir, "attach_pdf.scr");
            var sb = new StringBuilder(8192);
            var fmt = CultureInfo.InvariantCulture;

            foreach (var kv in map)
            {
                sb.AppendLine($".OPEN \"{kv.Key}\"");
                foreach (var d in kv.Value)
                {
                    sb.AppendLine($".-PDFATTACH \"PDF\\{d.FileName}\" {d.Page} " +
                                  $"{d.InsXmm.ToString(fmt)},{d.InsYmm.ToString(fmt)} " +
                                  $"{d.Scale.ToString(fmt)} {d.RotationDeg.ToString(fmt)}");
                }
                sb.AppendLine(".QSAVE");
            }
            sb.AppendLine(".QUIT");
            File.WriteAllText(scr, sb.ToString(), Encoding.ASCII);
            return scr;
        }

        public static void RunAccoreConsole(string exe, string scr)
        {
            var psi = new ProcessStartInfo(exe,
                           $"/s \"{scr}\" /isolate /l en-US")
            { UseShellExecute = false, CreateNoWindow = true };
            using var p = Process.Start(psi);
            p?.WaitForExit();
            if (p == null || p.ExitCode != 0)
                throw new InvalidOperationException("accoreconsole.exe lỗi hoặc không tồn tại.");
        }
    }
}
