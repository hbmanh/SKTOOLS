using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace EquipmentClassificationProcessor
{
    [Transaction(TransactionMode.Manual)]
    public class EquipmentClassificationProcessor : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            // Bước 1: Thu thập tất cả FamilyInstance
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType();

            // Bước 2: Đếm riêng theo từng Type trong từng nhóm Classification
            Dictionary<(string ecValue, ElementId typeId), int> ecTypeCounts = new Dictionary<(string, ElementId), int>();
            Dictionary<(string ecValue, ElementId typeId), FamilySymbol> symbolMap = new Dictionary<(string, ElementId), FamilySymbol>();

            foreach (FamilyInstance fi in collector)
            {
                FamilySymbol symbol = fi.Symbol;
                Parameter ecParam = symbol.LookupParameter("Equipment Classification-SP");
                if (ecParam == null || ecParam.StorageType != StorageType.String) continue;

                string ecValue = ecParam.AsString();
                if (string.IsNullOrWhiteSpace(ecValue)) continue;

                var key = (ecValue, symbol.Id);
                if (!ecTypeCounts.ContainsKey(key))
                    ecTypeCounts[key] = 0;

                ecTypeCounts[key]++;
                symbolMap[key] = symbol;
            }

            // Bước 3: Ghi kết quả đếm vào các Type Parameters
            using (Transaction tx = new Transaction(doc, "Update Equipment Quantities by Type and Classification"))
            {
                tx.Start();

                foreach (var kvp in ecTypeCounts)
                {
                    string ecValue = kvp.Key.ecValue;
                    ElementId typeId = kvp.Key.typeId;
                    int count = kvp.Value;
                    bool isCU = string.Equals(ecValue, "CU", StringComparison.OrdinalIgnoreCase);

                    if (!symbolMap.TryGetValue((ecValue, typeId), out FamilySymbol symbol))
                        continue;

                    Parameter indoor = symbol.LookupParameter("Q'ty Indoor Unit-SP");
                    Parameter outdoor = symbol.LookupParameter("Q'ty Outdoor Unit-SP");

                    if (isCU)
                    {
                        if (indoor != null && indoor.StorageType == StorageType.String)
                            indoor.Set("-");
                        if (outdoor != null && outdoor.StorageType == StorageType.String)
                            outdoor.Set(count.ToString());
                    }
                    else
                    {
                        if (indoor != null && indoor.StorageType == StorageType.String)
                            indoor.Set(count.ToString());
                        if (outdoor != null && outdoor.StorageType == StorageType.String)
                            outdoor.Set("-");
                    }
                }

                tx.Commit();
            }

            TaskDialog.Show("Hoàn tất", "Đã đếm và cập nhật số lượng theo từng Type trong mỗi nhóm Equipment Classification-SP.");
            return Result.Succeeded;
        }
    }
}
