using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.PointCloudAddins
{
    [Transaction(TransactionMode.Manual)]
    public class TogglePointCloudCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string msg, ElementSet set)
        {
            UIApplication uiApp = c.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            View view = doc.ActiveView;

            if (view == null)
            {
                TaskDialog.Show("PointCloud Toggle", "Không có view đang mở.");
                return Result.Failed;
            }

            // Nếu view đang dùng ViewTemplate → chỉ cảnh báo rồi thoát
            if (view.ViewTemplateId != ElementId.InvalidElementId)
            {
                string templateName = (doc.GetElement(view.ViewTemplateId) as View)?.Name ?? "Không xác định";
                TaskDialog.Show(
                    "PointCloud Toggle",
                    $"View “{view.Name}” đang dùng View Template “{templateName}”.\n" +
                    "Không thể bật/tắt PointCloud trực tiếp. Hãy chỉnh trong View Template nếu cần."
                );
                return Result.Cancelled;
            }

            // Nếu view KHÔNG dùng ViewTemplate → toggle PointCloud
            BuiltInCategory bic = BuiltInCategory.OST_PointClouds;
            ElementId catId = new ElementId(bic);

            using (Transaction t = new Transaction(doc, "Toggle PointCloud Visibility"))
            {
                try
                {
                    t.Start();

                    bool isHidden = view.GetCategoryHidden(catId);
                    view.SetCategoryHidden(catId, !isHidden);

                    t.Commit();
                }
                catch (Exception ex)
                {
                    msg = ex.Message;
                    if (t.HasStarted()) t.RollBack();
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }
    }
}
