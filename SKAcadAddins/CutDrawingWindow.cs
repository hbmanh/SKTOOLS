using System;
using System.IO;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SKAcadAddins.CutDrawingWindow))]

namespace SKAcadAddins
{
    public class CutDrawingWindow : IExtensionApplication
    {
        public void Initialize() { }
        public void Terminate() { }

        [CommandMethod("CDW")]
        public void ExecuteCutDrawingWindow()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Lấy thư mục của file DWG gốc
            string sourcePath = doc.Name;
            string folderPath = Path.GetDirectoryName(sourcePath);
            if (string.IsNullOrEmpty(folderPath))
            {
                ed.WriteMessage("\nKhông tìm thấy thư mục bản vẽ gốc.");
                return;
            }

            int fileIndex = 1;

            while (true)
            {
                // Chọn khung tên (Block hoặc Polyline)
                PromptEntityOptions peo = new PromptEntityOptions("\nChọn đối tượng khung tên (Block hoặc Polyline):");
                peo.SetRejectMessage("\nChỉ chọn Block hoặc Polyline.");
                peo.AddAllowedClass(typeof(BlockReference), true);
                peo.AddAllowedClass(typeof(Polyline), true);
                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                Point3d basePoint = Point3d.Origin;

                using (Transaction trNew = newDb.TransactionManager.StartTransaction())
                {
                    BlockTableRecord newMs = trNew.GetObject(newDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    // Mở khóa + mở layer
                    LayerTable lt = trNew.GetObject(newDb.LayerTableId, OpenMode.ForWrite) as LayerTable;
                    foreach (ObjectId layerId in lt)
                    {
                        LayerTableRecord ltr = trNew.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;

                        if (ltr.IsLocked)
                            ltr.IsLocked = false;

                        if (ltr.IsFrozen)
                            ltr.IsFrozen = false;

                        if (ltr.IsOff)
                            ltr.IsOff = false;
                    }

                    // Dịch toàn bộ đối tượng
                    foreach (ObjectId id in newMs)
                    {
                        Entity ent = trNew.GetObject(id, OpenMode.ForWrite) as Entity;
                        if (ent != null)
                        {
                            try { ent.TransformBy(Matrix3d.Displacement(moveVec)); }
                            catch { ed.WriteMessage($"\nKhông thể dịch đối tượng: {ent.GetType().Name}"); }
                        }
                    }

                    trNew.Commit();
                }


                PromptKeywordOptions pko = new PromptKeywordOptions("\nTiếp tục chọn bản vẽ khác? [Yes/No]");
                pko.Keywords.Add("Yes");
                pko.Keywords.Add("No");
                PromptResult pkr = ed.GetKeywords(pko);
                if (pkr.Status != PromptStatus.OK || pkr.StringResult == "No") break;
            }
        }

        /// <summary>
        /// Copy các Xref liên quan đến cùng thư mục nếu chưa tồn tại
        /// </summary>
        private void CopyXrefsIfNeeded(Database db, string targetFolder, Editor ed)
        {
            XrefGraph xg = db.GetHostDwgXrefGraph(false);
            for (int i = 1; i < xg.NumNodes; i++) // Bỏ qua node 0 (host)
            {
                XrefGraphNode node = xg.GetXrefNode(i) as XrefGraphNode;
                if (node.XrefStatus == XrefStatus.Resolved)
                {
                    string xrefPath = node.Database.Filename;
                    if (File.Exists(xrefPath))
                    {
                        string xrefName = Path.GetFileName(xrefPath);
                        string destPath = Path.Combine(targetFolder, xrefName);
                        if (!File.Exists(destPath))
                        {
                            try
                            {
                                File.Copy(xrefPath, destPath);
                                ed.WriteMessage($"\nĐã copy Xref: {xrefName}");
                            }
                            catch (Exception ex)
                            {
                                ed.WriteMessage($"\nKhông thể copy Xref: {xrefName} → {ex.Message}");
                            }
                        }
                    }
                }
            }
        }
    }
}
