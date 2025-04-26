using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(SKAcadAddins.GroupLayerCmd))]

namespace SKAcadAddins
{
    public class GroupLayerCmd : IExtensionApplication
    {
        public void Initialize() { }
        public void Terminate() { }

        [CommandMethod("SKGROUPLAYER")]
        public void GroupLayer()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForWrite);

                PromptStringOptions opts = new PromptStringOptions("\nNhập kí hiệu cho Layer: ");
                opts.AllowSpaces = true;
                PromptResult prefixResult = ed.GetString(opts);
                if (prefixResult.Status != PromptStatus.OK) return;
                string prefixCode = prefixResult.StringResult;

                // ✅ Duyệt trực tiếp và xử lý luôn
                ProcessEntities(ms, trans, lt, prefixCode);

                trans.Commit();
                ed.WriteMessage("\nĐã hoàn thành việc gộp và chuyển đổi Layer, vui lòng kiểm tra lại!");
            }
        }

        private void ProcessEntities(BlockTableRecord btr, Transaction trans, LayerTable lt, string prefixCode)
        {
            foreach (ObjectId id in btr)
            {
                Entity ent = trans.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                if (ent is BlockReference br)
                {
                    // ✅ Xử lý chính bản thân BlockReference
                    HandleEntityLayer(br, trans, lt, prefixCode);

                    // ✅ Xử lý AttributeReference nếu có
                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        if (!attId.IsErased)
                        {
                            AttributeReference attRef = trans.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                            if (attRef != null)
                            {
                                HandleEntityLayer(attRef, trans, lt, prefixCode);
                            }
                        }
                    }

                    // ✅ Đệ quy các đối tượng trong BlockReference
                    BlockTableRecord nestedBtr = (BlockTableRecord)trans.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                    ProcessEntities(nestedBtr, trans, lt, prefixCode);
                }
                else
                {
                    if (IsSupportedEntity(ent))
                    {
                        HandleEntityLayer(ent, trans, lt, prefixCode);
                    }
                }
            }
        }

        private bool IsSupportedEntity(Entity ent)
        {
            return ent is Line || ent is Arc || ent is Circle || ent is Polyline ||
                   ent is DBText || ent is MText || ent is Ellipse ||
                   ent is Hatch || ent is AttributeReference;
        }

        private void HandleEntityLayer(Entity entity, Transaction trans, LayerTable lt, string prefixCode)
        {
            var color = entity.Color;
            var linetypeId = entity.LinetypeId;
            var lineWeight = entity.LineWeight;

            string newLayerLinetypeName = entity.Linetype;
            var newLayerLinetypeId = entity.LinetypeId;

            if (newLayerLinetypeName.Equals("BYLAYER", StringComparison.OrdinalIgnoreCase))
            {
                LayerTableRecord layerRecord = trans.GetObject(entity.LayerId, OpenMode.ForRead) as LayerTableRecord;
                if (layerRecord != null)
                {
                    LinetypeTableRecord linetypeRecord = trans.GetObject(layerRecord.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;
                    if (linetypeRecord != null)
                    {
                        newLayerLinetypeName = linetypeRecord.Name;
                        newLayerLinetypeId = linetypeRecord.Id;
                    }
                }
            }

            if (lineWeight == LineWeight.ByLayer)
            {
                LayerTableRecord layerRecord = trans.GetObject(entity.LayerId, OpenMode.ForRead) as LayerTableRecord;
                if (layerRecord != null)
                {
                    lineWeight = layerRecord.LineWeight;
                }
            }

            string newLayerColor = entity.Color.ToString().Replace(",", "-");
            var colorR = color.ColorValue.R;
            var colorG = color.ColorValue.G;
            var colorB = color.ColorValue.B;
            if (newLayerColor.Equals("BYLAYER", StringComparison.OrdinalIgnoreCase))
            {
                LayerTableRecord layerRecord = trans.GetObject(entity.LayerId, OpenMode.ForRead) as LayerTableRecord;
                if (layerRecord != null)
                {
                    colorR = layerRecord.Color.ColorValue.R;
                    colorG = layerRecord.Color.ColorValue.G;
                    colorB = layerRecord.Color.ColorValue.B;
                    newLayerColor = $"{colorR}-{colorG}-{colorB}";
                }
            }

            string newLayerName = $"{prefixCode}_{newLayerLinetypeName}_{lineWeight}_{newLayerColor}";

            LayerTableRecord ltr = null;

            if (!lt.Has(newLayerName))
            {
                ltr = new LayerTableRecord
                {
                    Name = newLayerName,
                    Color = Color.FromRgb((byte)colorR, (byte)colorG, (byte)colorB),
                    LineWeight = lineWeight,
                    LinetypeObjectId = newLayerLinetypeId
                };
                lt.UpgradeOpen();
                lt.Add(ltr);
                trans.AddNewlyCreatedDBObject(ltr, true);
            }
            else
            {
                ltr = (LayerTableRecord)trans.GetObject(lt[newLayerName], OpenMode.ForRead);
            }

            if (ltr != null)
            {
                entity.UpgradeOpen();
                entity.LayerId = ltr.ObjectId;
            }
        }
    }
}
