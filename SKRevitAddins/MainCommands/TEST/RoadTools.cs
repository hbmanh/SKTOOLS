using System;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.PointCloudTools
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CheckPointCloudElevationCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string msg, ElementSet set)
        {
            UIApplication uiApp = c.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Lọc toàn bộ Point Cloud Instance trong project
                var pcs = new FilteredElementCollector(doc)
                    .OfClass(typeof(PointCloudInstance))
                    .Cast<PointCloudInstance>()
                    .ToList();

                if (pcs.Count == 0)
                {
                    TaskDialog.Show("PointCloud Elevation", "Không có PointCloud nào được link trong dự án.");
                    return Result.Succeeded;
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("KẾT QUẢ KIỂM TRA CAO ĐỘ POINTCLOUD (mm)");
                sb.AppendLine("============================================\n");

                foreach (var pc in pcs)
                {
                    string name = pc.Name;
                    Transform tf = pc.GetTotalTransform();
                    XYZ origin = tf.Origin;

                    double x_mm = UnitUtils.ConvertFromInternalUnits(origin.X, UnitTypeId.Millimeters);
                    double y_mm = UnitUtils.ConvertFromInternalUnits(origin.Y, UnitTypeId.Millimeters);
                    double z_mm = UnitUtils.ConvertFromInternalUnits(origin.Z, UnitTypeId.Millimeters);

                    sb.AppendLine($"Tên PointCloud: {name}");
                    sb.AppendLine($"- X: {x_mm:F1} mm");
                    sb.AppendLine($"- Y: {y_mm:F1} mm");
                    sb.AppendLine($"- Z (cao độ): {z_mm:F1} mm");
                    sb.AppendLine("--------------------------------------------");
                }

                TaskDialog.Show("PointCloud Elevation", sb.ToString());
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                TaskDialog.Show("Lỗi", $"Đã xảy ra lỗi: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}
