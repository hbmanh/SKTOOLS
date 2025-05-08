using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

[assembly: CommandClass(typeof(CadAddin.PrintAreaSetup))]

namespace CadAddin
{
    public class PrintAreaSetup
    {
        [CommandMethod("PrintArea")]
        [CommandMethod("PA")]
        public void RunPrintAreaSetup()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // --- 1. Chọn khổ giấy ---
            Dictionary<string, string> paperSizes = new Dictionary<string, string>
            {
                { "A4", "ISO A4 (210.00 x 297.00 MM)" },
                { "A3", "ISO A3 (297.00 x 420.00 MM)" },
                { "A2", "ISO A2 (420.00 x 594.00 MM)" },
                { "A1", "ISO A1 (594.00 x 841.00 MM)" },
                { "A0", "ISO A0 (841.00 x 1189.00 MM)" }
            };

            PromptKeywordOptions pko = new PromptKeywordOptions("\nChọn khổ giấy:");
            foreach (var size in paperSizes.Keys) pko.Keywords.Add(size);
            pko.AllowNone = false;

            PromptResult pres = ed.GetKeywords(pko);
            if (pres.Status != PromptStatus.OK) return;

            string paperSize = paperSizes[pres.StringResult];
            string paperSizeShort = pres.StringResult; // Lưu lại định dạng ngắn (A4, A3, v.v.)

            // --- 2. Kiểm tra xem đang ở trong Layout không và lấy thông tin ---
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Lấy layout hiện tại
                    string layoutName = LayoutManager.Current.CurrentLayout;
                    ObjectId layoutId = LayoutManager.Current.GetLayoutId(layoutName);
                    Layout layout = tr.GetObject(layoutId, OpenMode.ForWrite) as Layout;

                    if (layout == null)
                    {
                        ed.WriteMessage("\n❌ Không thể tìm thấy layout hiện tại.");
                        return;
                    }

                    if (layout.ModelType)
                    {
                        ed.WriteMessage("\n⚠️ Chỉ áp dụng trong Paper Space, không phải Model Space.");
                        return;
                    }

                    // --- 3. Chọn vùng in trong Paper Space ---
                    ed.WriteMessage("\nHãy chắc chắn rằng bạn đang chọn điểm trong Paper Space (Layout), không phải Model Space");
                    ed.WriteMessage("\nNếu bạn đang thấy nội dung bản vẽ trong viewport, hãy nhấp vào khu vực bên ngoài viewport.");

                    // Thiết lập điểm đầu tiên
                    PromptPointOptions ppo = new PromptPointOptions("\nChọn điểm thứ nhất:");
                    ppo.AllowNone = false;
                    ppo.LimitsChecked = false; // Cho phép chọn bất kỳ điểm nào

                    PromptPointResult p1 = ed.GetPoint(ppo);
                    if (p1.Status != PromptStatus.OK) return;

                    // Thiết lập điểm thứ hai
                    PromptCornerOptions p2opts = new PromptCornerOptions("\nChọn điểm đối diện:", p1.Value);
                    p2opts.LimitsChecked = false; // Cho phép chọn bất kỳ điểm nào

                    PromptPointResult p2 = ed.GetCorner(p2opts);
                    if (p2.Status != PromptStatus.OK) return;

                    // Tạo vùng in từ hai điểm
                    Point2d minPt = new Point2d(Math.Min(p1.Value.X, p2.Value.X), Math.Min(p1.Value.Y, p2.Value.Y));
                    Point2d maxPt = new Point2d(Math.Max(p1.Value.X, p2.Value.X), Math.Max(p1.Value.Y, p2.Value.Y));

                    // Đảm bảo tọa độ không âm nếu có thể
                    if (minPt.X < 0 || minPt.Y < 0)
                    {
                        ed.WriteMessage("\n⚠️ Cảnh báo: Vùng in có tọa độ âm, điều này có thể gây lỗi.");
                        // Có thể điều chỉnh để đảm bảo tọa độ không âm
                        minPt = new Point2d(Math.Max(0, minPt.X), Math.Max(0, minPt.Y));
                    }

                    Extents2d windowArea = new Extents2d(minPt, maxPt);

                    // Debug info
                    ed.WriteMessage($"\nVùng đã chọn: ({minPt.X},{minPt.Y}) đến ({maxPt.X},{maxPt.Y})");

                    // --- 4. Kiểm tra lại xem vùng in có hợp lệ không ---
                    double width = maxPt.X - minPt.X;
                    double height = maxPt.Y - minPt.Y;

                    if (width < 1.0 || height < 1.0)
                    {
                        ed.WriteMessage("\n❌ Vùng in quá nhỏ. Hãy chọn vùng lớn hơn.");
                        return;
                    }

                    // --- 5. Sử dụng phương pháp đơn giản hơn để thiết lập vùng in ---
                    // Tạo PlotSettings mới từ layout hiện tại
                    PlotSettings ps = new PlotSettings(layout.ModelType);
                    ps.CopyFrom(layout);

                    // Lấy PlotSettingsValidator và thiết lập từng bước
                    PlotSettingsValidator psv = PlotSettingsValidator.Current;

                    try
                    {
                        // Thiết lập loại in và vùng in
                        psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                        psv.SetPlotWindowArea(ps, windowArea);

                        // Cập nhật các cài đặt khác
                        psv.SetPlotCentered(ps, true);
                        psv.SetPlotRotation(ps, PlotRotation.Degrees000);
                        psv.SetPlotPaperUnits(ps, PlotPaperUnit.Millimeters);
                        psv.SetUseStandardScale(ps, true);
                        psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);

                        // Cập nhật layout từ PlotSettings đã thiết lập
                        layout.CopyFrom(ps);
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception acex)
                    {
                        // Xử lý lỗi AutoCAD cụ thể
                        ed.WriteMessage($"\n❌ Lỗi AutoCAD: {acex.Message} (Mã lỗi: {acex.ErrorStatus})");

                        // Thử phương pháp thay thế nếu phương pháp đầu tiên thất bại
                        try
                        {
                            // Chỉ thiết lập các thông số cơ bản
                            layout.TabOrder = 1; // Một thiết lập đơn giản để kiểm tra xem layout có thể ghi không

                            // Hiển thị cảnh báo và hướng dẫn
                            ed.WriteMessage("\n⚠️ Không thể thiết lập vùng in tự động.");
                            ed.WriteMessage("\nHướng dẫn thủ công:");
                            ed.WriteMessage("\n1. Chọn tab Layout");
                            ed.WriteMessage("\n2. Nhấp chuột phải và chọn Page Setup Manager");
                            ed.WriteMessage("\n3. Chọn Modify");
                            ed.WriteMessage("\n4. Trong tab Plot Area, chọn Window và nhập các tọa độ sau:");
                            ed.WriteMessage($"\n   - Lower left: ({minPt.X}, {minPt.Y})");
                            ed.WriteMessage($"\n   - Upper right: ({maxPt.X}, {maxPt.Y})");
                        }
                        catch (System.Exception ex2)
                        {
                            ed.WriteMessage($"\n❌ Lỗi bổ sung: {ex2.Message}");
                        }
                        return;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n❌ Lỗi khi thiết lập vùng in: {ex.Message}");
                        ed.WriteMessage($"\nChi tiết: {ex.ToString()}");
                        return;
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n✅ Đã thiết lập vùng in thành công cho khổ giấy {paperSizeShort}!");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                    return;
                }
            }
        }
    }
}