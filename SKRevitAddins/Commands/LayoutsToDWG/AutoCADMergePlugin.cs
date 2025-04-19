// AcadMergePlugin.AutoCADMergePlugin.cs
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.IO;

[assembly: CommandClass(typeof(AcadMergePlugin.MergeSheetsCommand))]
namespace AcadMergePlugin
{
    /// <summary>Plugin .NET nạp trong AutoCAD, nhận JSON & gộp DWG.</summary>
    public class MergeSheetsCommand
    {
        class CmdCfg { public string[] SheetFiles; public string FilePath; public bool MergeLayers; public bool OpenFile2; }

        [CommandMethod("MLabsMergeSheets")]
        public void MergeSheets()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            string json = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "merge_command.json");
            if (!File.Exists(json)) { ed.WriteMessage("\nJSON not found."); return; }

            CmdCfg cfg = JsonConvert.DeserializeObject<CmdCfg>(File.ReadAllText(json));
            if (cfg?.SheetFiles?.Length < 1) { ed.WriteMessage("\nInvalid DWG list."); return; }

            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                // mở file gốc
                doc.SendStringToExecute($"_.OPEN \"{cfg.SheetFiles[0]}\" ", true, false, false);

                // chèn các file còn lại
                for (int i = 1; i < cfg.SheetFiles.Length; i++)
                    doc.SendStringToExecute($"_.-INSERT \"{cfg.SheetFiles[i]}\" 0,0,0 1 0 ", true, false, false);

                // lưu DWG mới
                doc.SendStringToExecute($"_.SAVEAS 2013 \"{cfg.FilePath}\" ", true, false, false);
                if (!cfg.OpenFile2) doc.SendStringToExecute("_.QUIT ", true, false, false);
            }
            catch (System.Exception ex) { ed.WriteMessage($"\nMerge failed: {ex.Message}"); }
        }
    }
}