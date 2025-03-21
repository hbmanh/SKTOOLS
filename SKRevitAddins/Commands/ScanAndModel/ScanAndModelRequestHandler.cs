using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Linq;
using ScanAndModel.ViewModel;

namespace ScanAndModel
{
    public class ScanAndModelRequestHandler : IExternalEventHandler
    {
        private ScanAndModelViewModel _vm;
        private ScanAndModelRequest _request;

        public ScanAndModelRequestHandler(
            ScanAndModelViewModel viewModel,
            ScanAndModelRequest request)
        {
            _vm = viewModel;
            _request = request;
        }

        // Cho phép code-behind gọi handler.Request.Make(...)
        public ScanAndModelRequest Request => _request;

        public string GetName() => "ScanAndModelRequestHandler";

        public void Execute(UIApplication uiApp)
        {
            try
            {
                var reqId = _request.Take();
                switch (reqId)
                {
                    case ScanAndModelRequestId.None:
                        break;
                    case ScanAndModelRequestId.AutoDetectAndModel:
                        DoAutoDetectAndModel(uiApp);
                        break;
                    case ScanAndModelRequestId.ZoomToPoint:
                        DoZoomToPoint(uiApp);
                        break;
                }
            }
            catch (Exception ex)
            {
                _vm.StatusMessage = $"Error: {ex.Message}";
            }
        }

        private void DoAutoDetectAndModel(UIApplication uiApp)
        {
            var doc = uiApp.ActiveUIDocument.Document;

            using (Transaction trans = new Transaction(doc, "Auto Detect & Model"))
            {
                trans.Start();

                // 1) Tìm point cloud (placeholder)
                // var cloudInst = FindPointCloudInstance(doc);
                // if (cloudInst == null) { _vm.StatusMessage = "No point cloud found."; trans.RollBack(); return; }

                // 2) Giả lập detect => Tạo 1 bức tường
                var level = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .FirstOrDefault();
                if (level == null)
                {
                    _vm.StatusMessage = "No level found. Cannot create wall.";
                    trans.RollBack();
                    return;
                }

                // Tạo line placeholder
                XYZ p1 = new XYZ(0, 0, 0);
                XYZ p2 = new XYZ(10, 0, 0);
                Line wallLine = Line.CreateBound(p1, p2);

                Wall newWall = Wall.Create(doc, wallLine, level.Id, false);
                newWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(ElementId.InvalidElementId);
                newWall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(4);

                trans.Commit();
            }

            _vm.StatusMessage = "Auto-detect & model completed!";
        }

        private void DoZoomToPoint(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            View3D view3d = doc.ActiveView as View3D;
            if (view3d == null)
            {
                _vm.StatusMessage = "Active view is not 3D. Cannot zoom.";
                return;
            }

            // Giả sử zoom đến (5,5,0)
            XYZ target = new XYZ(5, 5, 0);
            double offset = 5;

            // Tạo bounding box => set SectionBox
            BoundingBoxXYZ bbox = new BoundingBoxXYZ
            {
                Min = target - new XYZ(offset, offset, offset),
                Max = target + new XYZ(offset, offset, offset)
            };

            using (Transaction tx = new Transaction(doc, "Zoom to point"))
            {
                tx.Start();
                view3d.SetSectionBox(bbox);
                tx.Commit();
            }

            uidoc.ActiveView = view3d;
            _vm.StatusMessage = "Zoomed to point cloud area. Please model manually.";
        }

        // private PointCloudInstance FindPointCloudInstance(Document doc) { ... }
    }
}
