using System;
using System.IO;
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

        [CommandMethod("CDW2")]
        public void CutMultipleFramesWithXrefCleanup()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            string folder = Path.GetDirectoryName(doc.Name);
            string baseName = Path.GetFileNameWithoutExtension(doc.Name);
            int fileIndex = 1;

            // --- BIND TOÀN BỘ XREF ---
            ObjectIdCollection xrefIds = new ObjectIdCollection();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (btr.IsFromExternalReference)
                    {
                        xrefIds.Add(btrId);
                    }
                }

                if (xrefIds.Count > 0)
                {
                    db.BindXrefs(xrefIds, true);
                }

                tr.Commit();
            }

            while (true)
            {
                PromptEntityOptions peo = new PromptEntityOptions("\nChọn khung tên (Xref hoặc Block):");
                peo.SetRejectMessage("\nChỉ được chọn Block hoặc Polyline.");
                peo.AddAllowedClass(typeof(BlockReference), true);
                peo.AddAllowedClass(typeof(Polyline), true);
                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) break;

                PromptPointResult pBase = ed.GetPoint("\nChọn điểm gốc mới (toạ độ 0,0):");
                if (pBase.Status != PromptStatus.OK) break;

                Extents3d bounds;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Entity ent = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                    bounds = ent.GeometricExtents;
                    tr.Commit();
                }

                using (Database newDb = new Database(true, true))
                {
                    // Giữ nguyên tỉ lệ LTS (line type scale)
                    newDb.Ltscale = db.Ltscale;

                    db.WblockCloneObjects(
                        GetAllModelSpaceIds(db),
                        newDb.CurrentSpaceId,
                        new IdMapping(),
                        DuplicateRecordCloning.Replace,
                        false
                    );

                    using (Transaction tr = newDb.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(newDb.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        LayerTable lt = (LayerTable)tr.GetObject(newDb.LayerTableId, OpenMode.ForWrite);
                        foreach (ObjectId lid in lt)
                        {
                            LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(lid, OpenMode.ForWrite);
                            ltr.IsLocked = false;
                            ltr.IsFrozen = false;
                            ltr.IsOff = false;
                        }

                        foreach (ObjectId id in ms)
                        {
                            Entity obj = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                            if (obj == null) continue;

                            // Bỏ qua BlockReference nằm bên trong khung
                            if (obj is BlockReference blockRef)
                            {
                                Extents3d? blkExt = TryGetGeometricExtents(blockRef);
                                if (blkExt.HasValue && IsInside(bounds, blkExt.Value))
                                    continue; // Không xoá nếu Block nằm trong khung
                            }

                            try
                            {
                                if (obj is BlockReference br)
                                {
                                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                                    if (btr.IsFromExternalReference && !IsInside(bounds, br.GeometricExtents))
                                    {
                                        br.Erase();
                                        continue;
                                    }
                                }

                                if (!IsInside(bounds, obj.GeometricExtents))
                                {
                                    obj.Erase();
                                }
                            }
                            catch
                            {
                                obj.Erase();
                            }
                        }

                        Vector3d moveVec = Point3d.Origin - pBase.Value;
                        foreach (ObjectId id in ms)
                        {
                            Entity ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                            ent?.TransformBy(Matrix3d.Displacement(moveVec));
                        }

                        tr.Commit();
                    }

                    string newFile = Path.Combine(folder, $"{baseName}_{fileIndex}.dwg");
                    newDb.SaveAs(newFile, DwgVersion.Current);
                    ed.WriteMessage($"\n→ Đã lưu: {newFile}");
                    fileIndex++;
                }

                PromptKeywordOptions pko = new PromptKeywordOptions("\nTiếp tục chọn khung khác? [Yes/No]", "Yes No");
                PromptResult pkr = ed.GetKeywords(pko);
                if (pkr.Status != PromptStatus.OK || pkr.StringResult == "No") break;
            }
        }

        private ObjectIdCollection GetAllModelSpaceIds(Database db)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                foreach (ObjectId id in ms)
                {
                    ids.Add(id);
                }
                tr.Commit();
            }
            return ids;
        }

        private bool IsInside(Extents3d outer, Extents3d inner)
        {
            return outer.MinPoint.X <= inner.MinPoint.X && outer.MaxPoint.X >= inner.MaxPoint.X &&
                   outer.MinPoint.Y <= inner.MinPoint.Y && outer.MaxPoint.Y >= inner.MaxPoint.Y;
        }

        private Extents3d? TryGetGeometricExtents(Entity ent)
        {
            try
            {
                return ent.GeometricExtents;
            }
            catch
            {
                return null;
            }
        }
    }
}
