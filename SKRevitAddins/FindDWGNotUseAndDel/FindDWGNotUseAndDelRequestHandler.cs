using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;
using SKRevitAddins.ViewModel;
namespace SKRevitAddins.FindDWGNotUseAndDel
{
    public class FindDWGNotUseAndDelRequestHandler : IExternalEventHandler
    {
        private FindDWGNotUseAndDelViewModel _vm;
        public FindDWGNotUseAndDelRequestHandler(FindDWGNotUseAndDelViewModel viewModel)
        {
            _vm = viewModel;
        }
        private FindDWGNotUseAndDelRequest m_Request = new FindDWGNotUseAndDelRequest();

        public FindDWGNotUseAndDelRequest Request
        {
            get { return m_Request; }
        }

        public void Execute(UIApplication uiApp)
        {
            try
            {
                Document doc = uiApp.ActiveUIDocument.Document;
                var reqId = Request.Take();
                switch (reqId)
                {
                    case RequestId.None: break;
                    case RequestId.Delete: DoDelete(doc); break;
                    case RequestId.OpenView: DoOpenView(doc); break;
                    case RequestId.Export: DoExport(); break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }
        public string GetName()
        {
            return "";
        }
        private void DoDelete(Document doc)
        {
            using (Transaction tx = new Transaction(doc, "Delete DWGs"))
            {
                tx.Start();
                if (_vm.IsDeleteAll)
                {
                    // Xoá All
                    var all = _vm.ImportedDWGs.ToList();
                    _vm.ImportedDWGs.Clear();

                    foreach (var dwg in all)
                    {
                        if (int.TryParse(dwg.InstanceId, out int idVal))
                        {
                            doc.Delete(new ElementId(idVal));
                        }
                    }
                }
                else
                {
                    // Xóa Selected
                    // 1. Lấy danh sách dòng đang chọn
                    var selected = _vm.SelectedDWGs?.ToList();
                    if (selected == null || selected.Count == 0)
                    {
                        TaskDialog.Show("Delete", "Không có DWG nào được chọn để xóa.");
                        tx.RollBack();
                        return;
                    }
                    // 2. Xóa từng item khỏi ViewModel và Revit
                    foreach (var dwg in selected)
                    {
                        _vm.ImportedDWGs.Remove(dwg);
                        if (int.TryParse(dwg.InstanceId, out int idVal))
                        {
                            doc.Delete(new ElementId(idVal));
                        }
                    }
                }
                tx.Commit();
            }
        }


        private void DoOpenView(Document doc)
        {
            // Giả sử ta yêu cầu người dùng chọn đúng 1 dòng
            if (_vm.SelectedDWGs == null || _vm.SelectedDWGs.Count != 1)
            {
                TaskDialog.Show("Open View", "Vui lòng chọn đúng 1 DWG để mở View.");
                return;
            }

            // Lấy item
            var dwgItem = _vm.SelectedDWGs[0];

            // Giả sử 'OwnerView' là tên View
            string viewName = dwgItem.OwnerView;
            if (string.IsNullOrEmpty(viewName))
            {
                TaskDialog.Show("Open View", "DWG này không có OwnerView hoặc tên View rỗng.");
                return;
            }

            // Tìm View trong Document theo tên
            var foundView = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .FirstOrDefault(v => v.Name == viewName);

            if (foundView == null)
            {
                TaskDialog.Show("Open View", $"Không tìm thấy View có tên: {viewName}");
                return;
            }

            // Mở View: cần UIDocument để set ActiveView
            var uiDoc = new UIDocument(doc);
            uiDoc.ActiveView = foundView;

            // (Tùy ý) Thông báo
            //TaskDialog.Show("Open View", $"Đã chuyển sang View: {foundView.Name}");
        }


        private void DoExport()
        {
            // 1. Tạo nội dung CSV
            var lines = _vm.ImportedDWGs.Select(d =>
                $"{d.InstanceId},{d.TypeId},{d.FileName},{d.InsertType},{d.Pinned},{d.OwnerView},{d.Group}");
            var csv = string.Join("\n", lines);

            // 2. Mở SaveFileDialog (sử dụng Microsoft.Win32.SaveFileDialog – có sẵn trong WPF)
            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Lưu bảng DWG xuất ra CSV",
                FileName = "ExportDWGs.csv",              // Tên file mặc định
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",  // Lọc file
                DefaultExt = ".csv"
            };

            bool? result = saveDialog.ShowDialog();  // Mở hộp thoại

            // 3. Nếu người dùng chọn xong
            if (result == true)
            {
                string filePath = saveDialog.FileName;   // Đường dẫn người dùng chọn
                try
                {
                    // 4. Ghi file CSV
                    System.IO.File.WriteAllText(filePath, csv);

                    // 5. Báo thành công
                    TaskDialog.Show("Export Table", $"Đã lưu file CSV:\n{filePath}");
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Export Error", $"Lỗi khi ghi file CSV: {ex.Message}");
                }
            }
            else
            {
                // Người dùng bấm Cancel hoặc đóng hộp thoại
                TaskDialog.Show("Export Table", "Đã hủy lưu file CSV.");
            }
        }

    }
}
