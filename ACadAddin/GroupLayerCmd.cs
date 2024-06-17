using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace ACadAddin
{
    public class GroupLayerCmd : IExtensionApplication
    {
        public void Initialize()
        {
            //throw new NotImplementedException();
        }

        public void Terminate()
        {
            //throw new NotImplementedException();
        }

        [CommandMethod("SKGROUPLAYER")]
        public void GroupLayer()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            Database db = doc.Database;
            Transaction trans = db.TransactionManager.StartTransaction();
            using (trans)
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForWrite);
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                LinetypeTable ltt = (LinetypeTable)trans.GetObject(db.LinetypeTableId, OpenMode.ForWrite);
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForWrite);

                TypedValue[] tv = new TypedValue[1];
                tv.SetValue(new TypedValue((int)DxfCode.Start, "LINE,ARC,CIRCLE,LWPOLYLINE,TEXT,MTEXT,ELLIPSE,INSERT"), 0);
                SelectionFilter filter = new SelectionFilter(tv);
                PromptSelectionResult psr = ed.SelectAll(filter);
                // Yêu cầu người dùng nhập đoạn mã
                PromptStringOptions opts = new PromptStringOptions("\nNhập kí hiệu cho Layer: ");
                opts.AllowSpaces = true;
                PromptResult prefixResult = ed.GetString(opts);
                if (prefixResult.Status != PromptStatus.OK) return;
                string prefixCode = prefixResult.StringResult;
                foreach (SelectedObject so in psr.Value)
                {
                    Entity entity = (Entity)trans.GetObject(so.ObjectId, OpenMode.ForWrite);
                    if (entity == null) continue;

                    var layerName = entity.Layer;
                    var color = entity.Color;
                    var linetypeId = entity.LinetypeId;
                    var lineWeight = entity.LineWeight;

                    // Lấy giá trị của Linetype từ layer ban đầu nếu là "ByLayer"
                    string newLayerLinetypeName = entity.Linetype;
                    var newLayerLinetypeId = entity.LinetypeId;
                    if (newLayerLinetypeName == "BYLAYER" || newLayerLinetypeName == "ByLayer")
                    {
                        LayerTableRecord layerRecord = trans.GetObject(entity.LayerId, OpenMode.ForRead) as LayerTableRecord;
                        if (layerRecord != null)
                        {
                            LinetypeTableRecord linetypeRecord = trans.GetObject(layerRecord.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;
                            if (linetypeRecord != null)
                            {
                                newLayerLinetypeName = linetypeRecord.Name.ToString();
                                newLayerLinetypeId = linetypeRecord.Id;
                            }
                        }
                    }

                    // Lấy giá trị của LineWeight từ layer ban đầu nếu là ByLayer
                    if (lineWeight == LineWeight.ByLayer)
                    {
                        LayerTableRecord layerRecord = trans.GetObject(entity.LayerId, OpenMode.ForRead) as LayerTableRecord;
                        if (layerRecord != null)
                        {
                            lineWeight = layerRecord.LineWeight;
                        }
                    }

                    // Lấy giá trị của Color từ layer ban đầu nếu là ByLayer
                    string newLayerColor = entity.Color.ToString().Replace(",", "-");
                    var colorR = color.ColorValue.R;
                    var colorG = color.ColorValue.G;
                    var colorB = color.ColorValue.B;
                    if (newLayerColor == "BYLAYER" )
                    {
                        LayerTableRecord layerRecord = trans.GetObject(entity.LayerId, OpenMode.ForRead) as LayerTableRecord;
                        if (layerRecord != null)
                        {
                            // Set newLayerColor to the color values from the layerRecord
                            colorR = layerRecord.Color.ColorValue.R;
                            colorG = layerRecord.Color.ColorValue.G;
                            colorB = layerRecord.Color.ColorValue.B;
                            newLayerColor = $"{colorR}-{colorG}-{colorB}";
                        }
                    }

                    string newLayerName = $"{prefixCode}_{newLayerLinetypeName}_{lineWeight}_{newLayerColor}";
                    bool layerExist = false;
                    LayerTableRecord ltr = null;
                    foreach (var layerId in lt)
                    {
                        var layer = trans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                        if (layer == null) continue;

                        if (layer.Name != newLayerName) continue;
                        layerExist = true;
                        ltr = layer;
                        break;
                    }
                    if (layerExist == false)
                    {
                        ltr = new LayerTableRecord();
                        ltr.Name = newLayerName;
                        ltr.Color = Color.FromRgb((byte)colorR, (byte)colorG, (byte)colorB);
                        ltr.LineWeight = lineWeight;
                        ltr.LinetypeObjectId = newLayerLinetypeId;
                        lt.UpgradeOpen();
                        ObjectId id = lt.Add(ltr);
                        trans.AddNewlyCreatedDBObject(ltr, true);
                    }

                    // Thay đổi layer của entity thành layer mới tương ứng
                    if (ltr != null && entity != null)
                    {
                        entity.LayerId = ltr.ObjectId;
                    }
                }
                trans.Commit();
                // Hiển thị thông báo khi hoàn thành
                ed.WriteMessage("\nĐã hoàn thành việc gộp và chuyển đổi Layer, vui lòng kiểm tra lại!");
            }
        }
    }
}
