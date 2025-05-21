using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Linq;
using SKRevitAddins.Utils;

namespace SKRevitAddins.SleeveChecker
{
    public static class SleeveCheckerLogic
    {
        public static void Run(Document doc, SleeveCheckerViewModel vm, ObservableCollection<ErrorItem> errors, bool previewOnly, UIApplication uiapp)
        {
            // Bạn copy lại nghiệp vụ thực tế ở đây, ví dụ:
            var beams = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .WhereElementIsNotElementType()
                .OfType<FamilyInstance>()
                .ToList();

            var meps = new FilteredElementCollector(doc)
                .OfClass(typeof(MEPCurve))
                .WhereElementIsNotElementType()
                .Cast<MEPCurve>()
                .ToList();

            foreach (var mep in meps)
            {
                if (mep.Name.Contains("A"))
                {
                    AddError(errors, "Test", "-", mep.Name, "Lỗi ví dụ", mep.Id, uiapp);
                }
            }
        }

        public static void AddError(ObservableCollection<ErrorItem> errors, string type, string frame, string mep, string msg, ElementId targetId, UIApplication uiapp)
        {
            var error = new ErrorItem
            {
                Type = type,
                Frame = frame,
                MEP = mep,
                Message = msg,
                TargetElementId = targetId
            };
            error.ShowCommand = new RelayCommand(_ => ShowElement(uiapp, targetId));
            errors.Add(error);
        }

        public static void ShowElement(UIApplication uiapp, ElementId eid)
        {
            var uidoc = uiapp.ActiveUIDocument;
            if (eid == null || eid == ElementId.InvalidElementId)
            {
                System.Windows.MessageBox.Show("Không tìm thấy đối tượng trong mô hình.", "Thông báo");
                return;
            }
            uidoc.Selection.SetElementIds(new List<ElementId> { eid });
            uidoc.ShowElements(eid);
        }
    }
}
