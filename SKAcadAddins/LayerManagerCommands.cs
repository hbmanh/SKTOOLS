//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using Autodesk.AutoCAD.ApplicationServices;
//using Autodesk.AutoCAD.Colors;
//using Autodesk.AutoCAD.DatabaseServices;
//using Autodesk.AutoCAD.EditorInput;
//using Autodesk.AutoCAD.Runtime;
//using Autodesk.AutoCAD.Windows;
//using Excel = Microsoft.Office.Interop.Excel;

//[assembly: CommandClass(typeof(LayerManager.LayerManagerCommands))]

//namespace LayerManager
//{
//    public class LayerManagerCommands
//    {
//        [CommandMethod("EXPORTLAYERS")]
//        public void ExportLayers()
//        {
//            Document doc = Application.DocumentManager.MdiActiveDocument;
//            Database db = doc.Database;
//            Editor ed = doc.Editor;

//            try
//            {
//                // Tạo file dialog để chọn vị trí lưu file Excel
//                SaveFileDialog sfd = new SaveFileDialog("Chọn vị trí lưu file Excel", "LayerExport.xlsx", "xlsx", "ExportLayers", SaveFileDialog.SaveFileDialogFlags.DoNotAddToRecent);

//                if (sfd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
//                    return;

//                string fileName = sfd.Filename;

//                // Bắt đầu xuất layer ra Excel
//                ed.WriteMessage("\nĐang xuất layers ra Excel...");

//                // Lấy thông tin tất cả layer
//                List<LayerInfo> layers = new List<LayerInfo>();

//                using (Transaction tr = db.TransactionManager.StartTransaction())
//                {
//                    LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

//                    foreach (ObjectId layerId in lt)
//                    {
//                        LayerTableRecord layer = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;

//                        layers.Add(new LayerInfo
//                        {
//                            Name = layer.Name,
//                            Color = layer.Color,
//                            Linetype = layer.LinetypeName,
//                            Lineweight = layer.LineWeight,
//                            IsOn = !layer.IsOff,
//                            IsLocked = layer.IsLocked,
//                            IsPlottable = !layer.IsPlotted
//                        });
//                    }

//                    tr.Commit();
//                }

//                // Xuất ra Excel
//                Excel.Application excelApp = new Excel.Application();
//                excelApp.Visible = false;

//                Excel.Workbook workbook = excelApp.Workbooks.Add();
//                Excel.Worksheet worksheet = workbook.Sheets[1];

//                // Thiết lập header
//                worksheet.Cells[1, 1] = "Layer Name";
//                worksheet.Cells[1, 2] = "Color (RGB)";
//                worksheet.Cells[1, 3] = "Linetype";
//                worksheet.Cells[1, 4] = "Lineweight";
//                worksheet.Cells[1, 5] = "Is On";
//                worksheet.Cells[1, 6] = "Is Locked";
//                worksheet.Cells[1, 7] = "Is Plottable";

//                // Bold header
//                Excel.Range headerRange = worksheet.Range["A1:G1"];
//                headerRange.Font.Bold = true;

//                // Điền dữ liệu
//                for (int i = 0; i < layers.Count; i++)
//                {
//                    LayerInfo layer = layers[i];
//                    int row = i + 2;

//                    worksheet.Cells[row, 1] = layer.Name;

//                    // Chuyển đổi màu sang RGB
//                    Color color = layer.Color;
//                    string colorRgb = string.Format("{0},{1},{2}", color.ColorValue.R, color.ColorValue.G, color.ColorValue.B);
//                    worksheet.Cells[row, 2] = colorRgb;

//                    worksheet.Cells[row, 3] = layer.Linetype;
//                    worksheet.Cells[row, 4] = layer.Lineweight.ToString();
//                    worksheet.Cells[row, 5] = layer.IsOn;
//                    worksheet.Cells[row, 6] = layer.IsLocked;
//                    worksheet.Cells[row, 7] = layer.IsPlottable;
//                }

//                // Tự động điều chỉnh độ rộng cột
//                worksheet.Columns.AutoFit();

//                // Lưu workbook
//                workbook.SaveAs(fileName);
//                workbook.Close();
//                excelApp.Quit();

//                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

//                ed.WriteMessage("\nĐã xuất thành công {0} layers ra file: {1}", layers.Count, fileName);
//            }
//            catch (System.Exception ex)
//            {
//                ed.WriteMessage("\nLỗi: " + ex.Message);
//            }
//        }

//        [CommandMethod("IMPORTLAYERS")]
//        public void ImportLayers()
//        {
//            Document doc = Application.DocumentManager.MdiActiveDocument;
//            Database db = doc.Database;
//            Editor ed = doc.Editor;

//            try
//            {
//                // Tạo file dialog để chọn file Excel
//                OpenFileDialog ofd = new OpenFileDialog("Chọn file Excel chứa dữ liệu layer", "xlsx", "xlsx", "ImportLayers", OpenFileDialog.OpenFileDialogFlags.DoNotAddToRecent);

//                if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
//                    return;

//                string fileName = ofd.Filename;

//                // Bắt đầu đọc từ Excel
//                ed.WriteMessage("\nĐang đọc dữ liệu layer từ Excel...");

//                Excel.Application excelApp = new Excel.Application();
//                excelApp.Visible = false;

//                Excel.Workbook workbook = excelApp.Workbooks.Open(fileName);
//                Excel.Worksheet worksheet = workbook.Sheets[1];

//                Excel.Range usedRange = worksheet.UsedRange;
//                int rowCount = usedRange.Rows.Count;

//                List<LayerInfo> layersToUpdate = new List<LayerInfo>();

//                // Đọc dữ liệu từ Excel (bắt đầu từ dòng 2, dòng 1 là header)
//                for (int i = 2; i <= rowCount; i++)
//                {
//                    string name = worksheet.Cells[i, 1].Value?.ToString();

//                    if (string.IsNullOrEmpty(name))
//                        continue;

//                    string colorRgb = worksheet.Cells[i, 2].Value?.ToString();
//                    string linetype = worksheet.Cells[i, 3].Value?.ToString();
//                    string lineweightStr = worksheet.Cells[i, 4].Value?.ToString();
//                    bool isOn = Convert.ToBoolean(worksheet.Cells[i, 5].Value);
//                    bool isLocked = Convert.ToBoolean(worksheet.Cells[i, 6].Value);
//                    bool isPlottable = Convert.ToBoolean(worksheet.Cells[i, 7].Value);

//                    // Phân tích chuỗi màu RGB
//                    Color color = Color.FromRgb(255, 255, 255); // Mặc định là trắng
//                    if (!string.IsNullOrEmpty(colorRgb))
//                    {
//                        string[] rgb = colorRgb.Split(',');
//                        if (rgb.Length >= 3)
//                        {
//                            byte r = Convert.ToByte(rgb[0]);
//                            byte g = Convert.ToByte(rgb[1]);
//                            byte b = Convert.ToByte(rgb[2]);
//                            color = Color.FromRgb(r, g, b);
//                        }
//                    }

//                    // Phân tích lineweight
//                    LineWeight lineweight = LineWeight.ByLineWeightDefault;
//                    if (!string.IsNullOrEmpty(lineweightStr))
//                    {
//                        if (Enum.TryParse(lineweightStr, out LineWeight lw))
//                            lineweight = lw;
//                    }

//                    layersToUpdate.Add(new LayerInfo
//                    {
//                        Name = name,
//                        Color = color,
//                        Linetype = linetype,
//                        Lineweight = lineweight,
//                        IsOn = isOn,
//                        IsLocked = isLocked,
//                        IsPlottable = isPlottable
//                    });
//                }

//                // Đóng Excel
//                workbook.Close(false);
//                excelApp.Quit();

//                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

//                // Cập nhật layers trong AutoCAD
//                using (Transaction tr = db.TransactionManager.StartTransaction())
//                {
//                    LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
//                    LinetypeTable ltt = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

//                    foreach (LayerInfo layerInfo in layersToUpdate)
//                    {
//                        // Kiểm tra xem layer đã tồn tại chưa
//                        if (lt.Has(layerInfo.Name))
//                        {
//                            // Cập nhật layer hiện có
//                            ObjectId layerId = lt[layerInfo.Name];
//                            LayerTableRecord layer = tr.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;

//                            layer.Color = layerInfo.Color;

//                            // Kiểm tra xem linetype có tồn tại không
//                            if (!string.IsNullOrEmpty(layerInfo.Linetype) && ltt.Has(layerInfo.Linetype))
//                                layer.LinetypeObjectId = ltt[layerInfo.Linetype];

//                            layer.LineWeight = layerInfo.Lineweight;
//                            layer.IsOff = !layerInfo.IsOn;
//                            layer.IsLocked = layerInfo.IsLocked;
//                            layer.IsPlotted = layerInfo.IsPlottable;
//                        }
//                        else
//                        {
//                            // Tạo layer mới
//                            LayerTableRecord newLayer = new LayerTableRecord();
//                            newLayer.Name = layerInfo.Name;
//                            newLayer.Color = layerInfo.Color;

//                            // Thiết lập linetype mặc định là "Continuous" nếu không tìm thấy
//                            if (!string.IsNullOrEmpty(layerInfo.Linetype) && ltt.Has(layerInfo.Linetype))
//                                newLayer.LinetypeObjectId = ltt[layerInfo.Linetype];
//                            else
//                                newLayer.LinetypeObjectId = ltt["Continuous"];

//                            newLayer.LineWeight = layerInfo.Lineweight;
//                            newLayer.IsOff = !layerInfo.IsOn;
//                            newLayer.IsLocked = layerInfo.IsLocked;
//                            newLayer.IsPlotted = layerInfo.IsPlottable;

//                            // Thêm layer mới vào bảng layer
//                            lt.UpgradeOpen();
//                            ObjectId newLayerId = lt.Add(newLayer);
//                            tr.AddNewlyCreatedDBObject(newLayer, true);
//                        }
//                    }

//                    tr.Commit();
//                }

//                ed.WriteMessage("\nĐã nhập thành công {0} layers.", layersToUpdate.Count);
//            }
//            catch (System.Exception ex)
//            {
//                ed.WriteMessage("\nLỗi: " + ex.Message);
//            }
//        }
//    }

//    // Lớp lưu trữ thông tin layer
//    public class LayerInfo
//    {
//        public string Name { get; set; }
//        public Color Color { get; set; }
//        public string Linetype { get; set; }
//        public LineWeight Lineweight { get; set; }
//        public bool IsOn { get; set; }
//        public bool IsLocked { get; set; }
//        public bool IsPlottable { get; set; }
//    }
//}