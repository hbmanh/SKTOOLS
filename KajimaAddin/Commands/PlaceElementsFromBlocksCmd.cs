using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace SKToolsAddins.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class PlaceElementsFromBlocksCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get the selected CAD link
            var refLinkCad = uidoc.Selection.PickObject(ObjectType.Element, new ImportInstanceSelectionFilter(), "Select Link File");
            var selectedCadLink = doc.GetElement(refLinkCad) as ImportInstance;
            Level level = uidoc.ActiveView.GenLevel;

            double offset = 0000 / 304.8;

            // Family and type names to place based on block names
            Dictionary<string, (BuiltInCategory category, string familyName, string typeName)> blockMappings = new Dictionary<string, (BuiltInCategory category, string familyName, string typeName)>
            {
                { "EA10168", (BuiltInCategory.OST_DuctTerminal, "033KM_線状吹出口_BL_SK", "線状吹出口_BL-D") },
                { "EA101B8", (BuiltInCategory.OST_DuctTerminal, "031_シーリングディフューザー_SK", "E2 #25") },
                { "EA10178", (BuiltInCategory.OST_DuctTerminal, "043KM_線状還気吸込口_CL_SK", "CL-4") },
                { "1D420270", (BuiltInCategory.OST_DuctTerminal, "041KM_ユニバーサル形排気吸込口_SK", "HS") },
                { "E640178", (BuiltInCategory.OST_DuctTerminal, "041KM_ユニバーサル形排気吸込口_SK", "HS") },
                { "E640188", (BuiltInCategory.OST_DuctTerminal, "020KM_天井ガラリ_SK", "排気ガラリ") },
                { "M800101", (BuiltInCategory.OST_DuctTerminal, "020KM_外気取入ガラリ_SK", "外気取入ガラリ") },
                { "EA301B8", (BuiltInCategory.OST_DuctTerminal, "020KM_排気ガラリ_SK", "排気ガラリ") },
                { "E610158", (BuiltInCategory.OST_DuctTerminal, "020KM_排気ガラリ_SK", "排気ガラリ") },
                { "E4301C8", (BuiltInCategory.OST_MechanicalEquipment, "07030_FCU-CID_天井埋込両ダクト形", "3CID") },
                { "1C860350", (BuiltInCategory.OST_MechanicalEquipment, "11030KM_FAN_消音ボックス付送風機", "#1_300m3/h") },
                { "E650178", (BuiltInCategory.OST_MechanicalEquipment, "07070KM_GHP_室外機", "112.0kW") },
                { "EA501C8", (BuiltInCategory.OST_MechanicalEquipment, "07070KM_GHP_室外機", "22.4kW") },
                { "E430128", (BuiltInCategory.OST_MechanicalEquipment, "07062KM_AI_室内機_カセット形(2方向)", "11.2kW") },
                { "E430198", (BuiltInCategory.OST_MechanicalEquipment, "07062KM_AI_室内機_カセット形(4方向)", "11.2kW") }
            };

            using (Transaction trans = new Transaction(doc, "Place Elements from CAD Blocks"))
            {
                trans.Start();

                // Get all block references from the CAD link
                Element elem = doc.GetElement(refLinkCad);
                GeometryElement geoElem = elem.get_Geometry(new Options());
                foreach (GeometryObject geoObj in geoElem)
                {
                    GeometryInstance instance = geoObj as GeometryInstance;
                    if (instance == null) continue;
                    foreach (GeometryObject instObj in instance.SymbolGeometry)
                    {
                        if (!(instObj is GeometryInstance blockInstance)) continue;
                        var blockName = blockInstance.Symbol.Name;

                        var blockInfo = blockMappings.FirstOrDefault(kv => blockName.Contains(kv.Key)).Value;
                        if (blockInfo == default) continue;
                        var (category, familyName, typeName) = blockInfo;

                        var familySymbol = new FilteredElementCollector(doc)
                            .OfCategory(category)
                            .OfClass(typeof(FamilySymbol))
                            .FirstOrDefault(e => (e as FamilySymbol).Family.Name == familyName && (e as FamilySymbol).Name == typeName) as FamilySymbol;

                        if (familySymbol == null) continue;
                        if (!familySymbol.IsActive)
                        {
                            familySymbol.Activate();
                            doc.Regenerate();
                        }

                        var blockPosition = blockInstance.Transform.Origin;
                        var blockRotation = blockInstance.Transform.BasisX.AngleTo(new XYZ(1, 0, 0));

                        XYZ placementPosition = new XYZ(blockPosition.X, blockPosition.Y, offset);
                        FamilyInstance familyInstance = doc.Create.NewFamilyInstance(placementPosition, familySymbol, level, StructuralType.NonStructural);

                        // Apply the rotation
                        ElementTransformUtils.RotateElement(doc, familyInstance.Id, Line.CreateBound(placementPosition, placementPosition + XYZ.BasisZ), blockRotation);
                    }
                }

                trans.Commit();
            }

            TaskDialog.Show("Success", "Successfully placed elements at block positions from the imported CAD file.");
            return Result.Succeeded;
        }

        class ImportInstanceSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is ImportInstance;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
    }
}
