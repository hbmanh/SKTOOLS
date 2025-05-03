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

            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType();

            // Bổ sung phân loại theo Classification, Equipment ID, và Level
            Dictionary<(string classification, string equipmentId, string levelName, ElementId typeId), int> ecTypeCounts =
                new Dictionary<(string, string, string, ElementId), int>();

            Dictionary<(string, string, string, ElementId typeId), FamilySymbol> symbolMap =
                new Dictionary<(string, string, string, ElementId), FamilySymbol>();

            foreach (FamilyInstance fi in collector)
            {
                FamilySymbol symbol = fi.Symbol;
                Parameter ecParam = symbol.LookupParameter("Equipment Classification-SP");
                Parameter eidParam = symbol.LookupParameter("Equipment ID-SP");
                Level level = doc.GetElement(fi.LevelId) as Level;

                if (ecParam == null || eidParam == null || level == null) continue;
                if (ecParam.StorageType != StorageType.String || eidParam.StorageType != StorageType.String) continue;

                string ecValue = ecParam.AsString();
                string eidValue = eidParam.AsString();
                string levelName = level.Name;

                if (string.IsNullOrWhiteSpace(ecValue) || string.IsNullOrWhiteSpace(eidValue)) continue;

                var key = (ecValue, eidValue, levelName, symbol.Id);
                if (!ecTypeCounts.ContainsKey(key))
                    ecTypeCounts[key] = 0;

                ecTypeCounts[key]++;
                symbolMap[key] = symbol;
            }

            using (Transaction tx = new Transaction(doc, "Update Equipment Quantities by Classification, ID, Level"))
            {
                tx.Start();

                foreach (var kvp in ecTypeCounts)
                {
                    string ecValue = kvp.Key.classification;
                    int count = kvp.Value;
                    ElementId typeId = kvp.Key.typeId;

                    if (!symbolMap.TryGetValue(kvp.Key, out FamilySymbol symbol))
                        continue;

                    Parameter indoor = symbol.LookupParameter("Q'ty Indoor Unit-SP");
                    Parameter outdoor = symbol.LookupParameter("Q'ty Outdoor Unit-SP");

                    bool isCU = string.Equals(ecValue, "CU", StringComparison.OrdinalIgnoreCase);

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

            TaskDialog.Show("Hoàn tất", "Đã đếm và cập nhật số lượng theo từng Classification, Equipment ID và Level.");
            return Result.Succeeded;
        }
    }
}
