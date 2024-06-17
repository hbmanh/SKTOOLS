using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit;
using Autodesk.Revit.DB;

namespace SKToolsAddins.Utils
{
    public static class ModelComparisonUtils
    {
        public class Diff
        {
            public string S1 { get; set; }
            public string S2 { get; set; }
            public string S3 { get; set; }
            public string S4 { get; set; }
            public Diff(string s1 = "", string s2 = "", string s3 = "", string s4 = "")
            {
                S1 = s1;
                S2 = s2;
                S3 = s3;
                S4 = s4;
            }
        }
        public class CatFamPar
        {
            public CatFamPar(string category, string family, string instOrType, string param, bool unchange = false)
            {
                Category = category;
                Family = family;
                InstOrType = instOrType;
                Parameter = param;
                Unchange = unchange;
            }
            public string Category { get; set; }
            public string Family { get; set; }
            public string InstOrType { get; set; }
            public string Parameter { get; set; }
            public bool Unchange { get; set; } = false;
        }
        public static (List<Element>, List<Element>, List<Element>, List<string>, List<Diff>, List<List<string>>) CompareModel (Document doc, Document linkedDoc, string option, List<Element> eles, List<Element> linkedEles, List<ElementId> typeEles, List<ElementId> linkedTypeEles, bool update = false, List<string> unchangedPars = null)
        {
            List<Element> newEles = new List<Element>();
            List<Element> modifiedEles = new List<Element>();
            List<Element> deletedEles = new List<Element>();
            List<Solid> deletedSolids = new List<Solid>();
            List<string> diffStrs = new List<string>();
            List<Diff> diffs = new List<Diff>();
            List<string> col1 = new List<string>();
            List<string> col2 = new List<string>();
            List<string> col3 = new List<string>();
            List<string> col4 = new List<string>();
            
            if (eles.Count > 0 && linkedEles.Count > 0)
            {
                foreach (var ele in eles)
                {
                    if (ele != null)
                    {
                        List<string> tempDiffStrs = new List<string>();
                        List<Diff> tempDiffs = new List<Diff>();
                        string eleName = ele.Name;
                        ElementId eleId = ele.Id;
                        string savedIdStr = null;
                        try
                        {
                            savedIdStr = ele.LookupParameter("リンクオブジェクト_Id").AsString();
                        }
                        catch
                        {
                        }
                        
                        int saveIdInt = 0;
                        if (savedIdStr != null)
                        {
                            saveIdInt = int.Parse(savedIdStr);
                        }
                        ElementId savedId = new ElementId(saveIdInt);
                        var linkedEle = linkedDoc.GetElement(eleId);
                        if (linkedEle is null)
                        {
                            linkedEle = linkedDoc.GetElement(savedId);
                        }
                        bool notMove = true;
                        if (linkedEle is null || linkedEle.Category is null)
                        {
                            switch (option)
                            {
                                case "リンク：新｜現在：古":
                                    deletedEles.Add(ele);
                                    break;
                                case "リンク：古｜現在：新":
                                    newEles.Add(ele);
                                    break;
                                default:
                                    break;
                            }
                        }
                        else
                        {
                            #region Location Checking
                            var loc = ele.Location;
                            var linkedLoc = linkedEle.Location;
                            if (loc is LocationPoint)
                            {
                                var locPoint = (loc as LocationPoint).Point;
                                var linkedLocPoint = (linkedLoc as LocationPoint).Point;
                                notMove = PointUtils.CheckCoincidentPoints(locPoint, linkedLocPoint, true);
                                if (update == true && notMove == false)
                                {
                                    ElementTransformUtils.MoveElement(doc, ele.Id, linkedLocPoint - locPoint);
                                }
                            }
                            else if (loc is LocationCurve)
                            {
                                var locCurve = (loc as LocationCurve).Curve;
                                var linkedLocCurve = (linkedLoc as LocationCurve).Curve;
                                notMove = locCurve.CheckTwoCurvesIfTheyAreSame(linkedLocCurve);
                                if (update == true && notMove == false)
                                {
                                    (ele.Location as LocationCurve).Curve = linkedLocCurve;
                                }
                            }
                            #endregion Location Checking

                            #region Geometry Checking
                            bool profileChanged = false;

                            List<Solid> solids = new List<Solid>();
                            solids = ele.GetAllSolidsAdvance(true);

                            // Get Solid if Element is Curtain Roof
                            if (eleName.Contains("地盤改良"))
                            {
                                var panels = new FilteredElementCollector(doc)
                                            .WhereElementIsNotElementType()
                                            .OfCategory(BuiltInCategory.OST_CurtainWallPanels)
                                            .OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                            .Where(e => e.Host.Id.IntegerValue == ele.Id.IntegerValue);
                                if (panels != null && panels.Count() > 0)
                                {
                                    var panel = panels.First();
                                    solids = panel.GetAllSolidsAdvance(true);
                                }
                            }

                            // Get Solid if Element is Floor
                            if (ele.Category.Id.IntegerValue == -2000032)
                            {
                                var floorCurveLoop = ele.GetFloorCurveLoop(doc);
                                var floorCurveLoops = new List<CurveLoop> { floorCurveLoop };
                                var floorThickness = (ele as Floor).FloorType.get_Parameter(BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM).AsDouble();
                                var floorExtrude = GeometryCreationUtilities.CreateExtrusionGeometry(floorCurveLoops, XYZ.BasisZ * -1, floorThickness);
                                solids = new List<Solid> { floorExtrude };
                            }

                            Solid solid = null;
                            if (solids != null)
                            {
                                solid = solids.UnionSolidList();
                            }

                            List<Solid> linkedSolids = linkedEle.GetAllSolidsAdvance(true);

                            // Get Solid if Linked Element is Curtain Roof
                            if (linkedEle.Name.Contains("地盤改良"))
                            {
                                var linkedPanels = new FilteredElementCollector(linkedDoc)
                                                .WhereElementIsNotElementType()
                                                .OfCategory(BuiltInCategory.OST_CurtainWallPanels)
                                                .OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                                .Where(e => e.Host.Id.IntegerValue == linkedEle.Id.IntegerValue);
                                if (linkedPanels != null && linkedPanels.Count() > 0)
                                {
                                    var linkedPanel = linkedPanels.First();
                                    linkedSolids = linkedPanel.GetAllSolidsAdvance(true);
                                }
                            }

                            // Get Solid if Linked Element is Floor
                            if (linkedEle.Category.Id.IntegerValue == -2000032)
                            {
                                var floorCurveLoop = linkedEle.GetFloorCurveLoop(linkedDoc);
                                var floorCurveLoops = new List<CurveLoop> { floorCurveLoop };
                                var floorThickness = (ele as Floor).FloorType.get_Parameter(BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM).AsDouble();
                                var floorExtrude = GeometryCreationUtilities.CreateExtrusionGeometry(floorCurveLoops, XYZ.BasisZ * -1, floorThickness);
                                linkedSolids = new List<Solid> { floorExtrude };
                            }

                            

                            Solid linkedSolid = null;
                            if (linkedSolids != null)
                            {
                                linkedSolid = linkedSolids.UnionSolidList();
                            }

                            if (solid != null && solid.Volume > 0
                                && linkedSolid != null && linkedSolid.Volume > 0)
                            {
                                try
                                {
                                    var xSolid1 = BooleanOperationsUtils.ExecuteBooleanOperation(solid, linkedSolid, BooleanOperationsType.Difference);
                                    var xSolid2 = BooleanOperationsUtils.ExecuteBooleanOperation(linkedSolid, solid, BooleanOperationsType.Difference);

                                    //var solidCentroid = solid.ComputeCentroid();
                                    //var linkedSolidCentroid = linkedSolid.ComputeCentroid();

                                    //notMove = PointUtils.CheckCoincidentPoints(solidCentroid, linkedSolidCentroid, true);

                                    if ((xSolid1 != null && xSolid1.Volume > 0)
                                        || (xSolid2 != null && xSolid2.Volume > 0))
                                    {
                                        modifiedEles.Add(ele);

                                        if (ele.Name.Contains("地盤改良")
                                            || (ele.Category.Id.IntegerValue == -2000032)
                                            || (ele.Category.Id.IntegerValue == -2000011))
                                        {
                                            profileChanged = true;
                                        }
                                        //if (ele.Category.Id.IntegerValue == -2000011)
                                        //{
                                        //    var wall = ele as Wall;
                                        //    var wallType = wall.WallType;
                                        //    var wallWidth = wallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble();
                                        //    var linkedWall = linkedEle as Wall;
                                        //    var linkedWallType = linkedWall.WallType;
                                        //    var linkedWallWidth = linkedWallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble();
                                        //    if (Math.Round(wallWidth - linkedWallWidth, 3) == 0)
                                        //    {
                                        //        profileChanged = true;
                                        //    }
                                        //}
                                    }
                                }
                                catch
                                {
                                }
                                
                            }

                            #endregion Geometry Checking

                            #region Copy New Instance if Profile changed
                            Element copiedEle = null;
                            if (profileChanged)
                            {
                                try
                                {
                                    List<ElementId> copiedEleIds = new List<ElementId>();
                                    using (SubTransaction subTxCopy = new SubTransaction(doc))
                                    {
                                        subTxCopy.Start();
                                        CopyPasteOptions copyPasteOptions = new CopyPasteOptions();
                                        copyPasteOptions.SetDuplicateTypeNamesHandler(new CopyUseDestination());
                                        List<ElementId> elesToCopy = new List<ElementId>() { linkedEle.Id };
                                        copiedEleIds = (List<ElementId>)ElementTransformUtils.CopyElements(linkedDoc, elesToCopy, doc, null, copyPasteOptions);
                                        deletedEles.Add(ele);

                                        var linkedEleDesignOption = linkedEle.DesignOption;

                                        subTxCopy.Commit();
                                    }

                                    var copiedEleId = copiedEleIds.First();
                                    copiedEle = doc.GetElement(copiedEleId);
                                    copiedEle.LookupParameter("リンクオブジェクト_Id").Set(linkedEle.Id.IntegerValue.ToString());
                                }
                                catch
                                {
                                }
                            }
                            #endregion Copy New Instance if Profile changed

                            #region Instance Parameters Checking
                            ParameterSet paras = ele.Parameters;
                            foreach (Parameter para in paras)
                            {
                                string diffStr = "";
                                Diff diff = new Diff();
                                bool paraHasValue = para.HasValue;
                                StorageType type = para.StorageType;
                                Parameter linkedPara = null;
                                var paraId = (para.Definition as InternalDefinition).BuiltInParameter;
                                try
                                {
                                    linkedPara = linkedEle.get_Parameter(paraId);
                                }
                                catch
                                {
                                }
                                if (linkedPara == null)
                                {
                                    linkedPara = linkedEle.get_Parameter(para.Definition);
                                }
                                if (linkedPara == null)
                                {
                                    linkedPara = linkedEle.LookupParameter(para.Definition.Name);
                                }
                                if (paraHasValue || (linkedPara != null && linkedPara.HasValue))
                                {
                                    switch (type)
                                    {
                                        case StorageType.None:
                                            break;
                                        case StorageType.Integer:
                                            var valueInt = para.AsInteger();
                                            if (linkedPara != null)
                                            {
                                                int linkedValueInt = linkedPara.AsInteger();
                                                if (valueInt != linkedValueInt)
                                                {
                                                    diffStr = para.Definition.Name + "," + linkedValueInt + "," + valueInt;
                                                    diff.S1 = para.Definition.Name;
                                                    diff.S2 = linkedValueInt.ToString();
                                                    diff.S3 = valueInt.ToString();

                                                    if (update && !para.IsReadOnly)
                                                    {
                                                        if (unchangedPars is null || (unchangedPars != null && !unchangedPars.Contains(para.Definition.Name)))
                                                        {
                                                            ele.LookupParameter(para.Definition.Name).Set(linkedValueInt);
                                                        }
                                                        else
                                                        {
                                                            copiedEle.LookupParameter(para.Definition.Name).Set(valueInt);
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                diffStr = para.Definition.Name + ",パラメーターを追加しました。";
                                                diff.S1 = para.Definition.Name;
                                                diff.S2 = "パラメーターを追加しました。";
                                            }
                                            break;
                                        case StorageType.Double:
                                            var valueDbl = Math.Round(para.AsDouble() * 304.8, 3);
                                            if (linkedPara != null)
                                            {
                                                double linkedValueDbl = Math.Round(linkedPara.AsDouble() * 304.8, 3);
                                                if (valueDbl != linkedValueDbl)
                                                {
                                                    diffStr = para.Definition.Name + "," + linkedValueDbl + "," + valueDbl;
                                                    diff.S1 = para.Definition.Name;
                                                    diff.S2 = linkedValueDbl.ToString();
                                                    diff.S3 = valueDbl.ToString();

                                                    if (update && !para.IsReadOnly)
                                                    {
                                                        if (unchangedPars is null || (unchangedPars != null && !unchangedPars.Contains(para.Definition.Name)))
                                                        {
                                                            ele.LookupParameter(para.Definition.Name).Set(linkedValueDbl / 304.8);
                                                        }
                                                        else
                                                        {
                                                            copiedEle.LookupParameter(para.Definition.Name).Set(valueDbl);
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                diffStr = para.Definition.Name + ",パラメーターを追加しました。";
                                                diff.S1 = para.Definition.Name;
                                                diff.S2 = "パラメーターを追加しました。";
                                            }
                                            break;
                                        case StorageType.String:
                                            var valueStr = para.AsString();
                                            if (linkedPara != null)
                                            {
                                                string linkedValueStr = linkedPara.AsString();
                                                if (valueStr != linkedValueStr)
                                                {
                                                    diffStr = para.Definition.Name + "," + linkedValueStr + "," + valueStr;
                                                    diff.S1 = para.Definition.Name;
                                                    diff.S2 = linkedValueStr;
                                                    diff.S3 = valueStr;

                                                    if (update && !para.IsReadOnly)
                                                    {
                                                        if (unchangedPars is null || (unchangedPars != null && !unchangedPars.Contains(para.Definition.Name)))
                                                        {
                                                            ele.LookupParameter(para.Definition.Name).Set(linkedValueStr);
                                                        }
                                                        else
                                                        {
                                                            copiedEle.LookupParameter(para.Definition.Name).Set(valueStr);
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                diffStr = para.Definition.Name + ",パラメーターを追加しました。";
                                                diff.S1 = para.Definition.Name;
                                                diff.S2 = "パラメーターを追加しました。";
                                            }
                                            break;
                                        case StorageType.ElementId:
                                            var valueEid = para.AsElementId();
                                            if (linkedPara != null)
                                            {
                                                ElementId linkedValueEid = linkedPara.AsElementId();
                                                if (valueEid != linkedValueEid)
                                                {
                                                    diffStr = para.Definition.Name + "," + linkedValueEid + "," + valueEid;
                                                    diff.S1 = para.Definition.Name;
                                                    diff.S2 = linkedValueEid.ToString();
                                                    diff.S3 = valueEid.ToString();

                                                    if (update && !para.IsReadOnly)
                                                    {
                                                        if (unchangedPars is null || (unchangedPars != null && !unchangedPars.Contains(para.Definition.Name)))
                                                        {
                                                            ele.LookupParameter(para.Definition.Name).Set(linkedValueEid);
                                                        }
                                                        else
                                                        {
                                                            copiedEle.LookupParameter(para.Definition.Name).Set(valueEid);
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                diffStr = para.Definition.Name + ",パラメーターを追加しました。";
                                                diff.S1 = para.Definition.Name;
                                                diff.S2 = "パラメーターを追加しました。";
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                }
                                if (diffStr != "")
                                {
                                    tempDiffStrs.Add(diffStr);
                                    tempDiffs.Add(diff);
                                }
                            }
                            #endregion Instance Parameters Checking
                        }
                        if (tempDiffStrs.Count > 0)
                        {
                            string eleDefStr = "INSTANCE," + ele.Category.Name + "," + ele.Name + "," + ele.Id;
                            Diff eleDef = new Diff("INSTANCE", ele.Category.Name, ele.Name, ele.Id.ToString());
                            diffStrs.Add(eleDefStr);
                            diffs.Add(eleDef);

                            if (notMove == false)
                            {
                                string diffStr = ele.Name + ",移動されました。";
                                Diff diff = new Diff();
                                diff.S1 = ele.Name;
                                diff.S2 = "移動されました。";
                                diffStrs.Add(diffStr);
                                diffs.Add(diff);
                            }

                            tempDiffStrs.ForEach(s => diffStrs.Add(s));
                            tempDiffs.ForEach(d => diffs.Add(d));
                            modifiedEles.Add(ele);
                        }
                    }
                }

                foreach (var linkedEle in linkedEles)
                {
                    ElementId linkedEleId = linkedEle.Id;
                    var ele = doc.GetElement(linkedEleId);
                    if (ele is null)
                    {
                        var elesGetByParam = new FilteredElementCollector(doc)
                                        .WhereElementIsNotElementType()
                                        .Where(e => (e.GetParameters("リンクオブジェクト_Id").Count > 0)
                                        && (e.LookupParameter("リンクオブジェクト_Id").AsString() == linkedEleId.ToString()));
                        if (elesGetByParam != null && elesGetByParam.Count() > 0)
                        {
                            ele = elesGetByParam.First();
                        }
                    }
                    if (linkedEle != null)
                    {
                        if (ele is null || ele.Category is null)
                        {
                            switch (option)
                            {
                                case "リンク：新｜現在：古":
                                    newEles.Add(linkedEle);
                                    break;
                                case "リンク：古｜現在：新":
                                    deletedEles.Add(linkedEle);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }

            if (typeEles.Count > 0 && linkedTypeEles.Count > 0)
            {
                foreach (ElementId typeEleId in typeEles)
                {
                    List<string> tempDiffStrs = new List<string>();
                    List<Diff> tempDiffs = new List<Diff>();
                    var typeEle = doc.GetElement(typeEleId);
                    if (typeEle != null)
                    {
                        string typeEleName = typeEle.Name;
                        var linkedTypeEle = linkedDoc.GetElement(typeEleId);
                        if (linkedTypeEle is null)
                        {
                            string newType = typeEleName + "," + typeEleId;
                            diffStrs.Add("New Type," + newType);
                            Diff diff = new Diff("New Type", newType, "", "");
                            diffs.Add(diff);
                        }
                        else
                        {
                            #region Type Parameters Checking
                            var typeParas = typeEle.Parameters;

                            foreach (Parameter typePara in typeParas)
                            {
                                string diffStr = "";
                                Diff diff = new Diff();
                                bool hasValue = typePara.HasValue;
                                StorageType type = typePara.StorageType;
                                Parameter linkedTypePara = null;
                                var typeParaId = (typePara.Definition as InternalDefinition).BuiltInParameter;
                                try
                                {
                                    linkedTypePara = linkedTypeEle.get_Parameter(typeParaId);
                                }
                                catch
                                {
                                }
                                if (linkedTypePara == null)
                                {
                                    linkedTypePara = linkedTypeEle.get_Parameter(typePara.Definition);
                                }
                                if (linkedTypePara == null)
                                {
                                    linkedTypePara = linkedTypeEle.LookupParameter(typePara.Definition.Name);
                                }
                                if (hasValue)
                                {
                                    switch (type)
                                    {
                                        case StorageType.None:
                                            break;
                                        case StorageType.Integer:
                                            var valueInt = typePara.AsInteger();
                                            if (linkedTypePara != null)
                                            {
                                                int linkedValueInt = linkedTypePara.AsInteger();
                                                if (valueInt != linkedValueInt)
                                                {
                                                    diffStr = typePara.Definition.Name + "," + linkedValueInt + "," + valueInt;
                                                    diff.S1 = typePara.Definition.Name;
                                                    diff.S2 = linkedValueInt.ToString();
                                                    diff.S3 = valueInt.ToString();

                                                    if (update && !typePara.IsReadOnly)
                                                    {
                                                        if (unchangedPars is null || (unchangedPars != null && !unchangedPars.Contains(typePara.Definition.Name)))
                                                        {
                                                            typeEle.LookupParameter(typePara.Definition.Name).Set(linkedValueInt);
                                                        } 
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                diffStr = typePara.Definition.Name + ",パラメーターを追加しました。";
                                                diff.S1 = typePara.Definition.Name;
                                                diff.S2 = "パラメーターを追加しました。";
                                            }
                                            break;
                                        case StorageType.Double:
                                            var valueDbl = Math.Round(typePara.AsDouble() * 304.8, 3);
                                            if (linkedTypePara != null)
                                            {
                                                double linkedValueDbl = Math.Round(linkedTypePara.AsDouble() * 304.8, 3);
                                                if (valueDbl != linkedValueDbl)
                                                {
                                                    diffStr = typePara.Definition.Name + "," + linkedValueDbl + "," + valueDbl;
                                                    diff.S1 = typePara.Definition.Name;
                                                    diff.S2 = linkedValueDbl.ToString();
                                                    diff.S3 = valueDbl.ToString();

                                                    if (update && !typePara.IsReadOnly)
                                                    {
                                                        if (unchangedPars is null || (unchangedPars != null && !unchangedPars.Contains(typePara.Definition.Name)))
                                                        {
                                                            typeEle.LookupParameter(typePara.Definition.Name).Set(linkedValueDbl / 304.8);
                                                        } 
                                                    }
                                                    if (update && typeEle.Category.Id.IntegerValue == -2000011 && typePara.Definition.Name == "Width")
                                                    {
                                                        CompoundStructure linkedCompoundStructure = (linkedTypeEle as WallType).GetCompoundStructure();
                                                        (typeEle as WallType).SetCompoundStructure(linkedCompoundStructure);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                diffStr = typePara.Definition.Name + ",パラメーターを追加しました。";
                                                diff.S1 = typePara.Definition.Name;
                                                diff.S2 = "パラメーターを追加しました。";
                                            }
                                            break;
                                        case StorageType.String:
                                            var valueStr = typePara.AsString();
                                            if (linkedTypePara != null)
                                            {
                                                string linkedValueStr = linkedTypePara.AsString();
                                                if (valueStr != linkedValueStr)
                                                {
                                                    diffStr = typePara.Definition.Name + "," + linkedValueStr + "," + valueStr;
                                                    diff.S1 = typePara.Definition.Name;
                                                    diff.S2 = linkedValueStr;
                                                    diff.S3 = valueStr;

                                                    if (update && !typePara.IsReadOnly)
                                                    {
                                                        if (unchangedPars is null || (unchangedPars != null && !unchangedPars.Contains(typePara.Definition.Name)))
                                                        {
                                                            typeEle.LookupParameter(typePara.Definition.Name).Set(linkedValueStr);
                                                        }
                                                    }
                                                    if (update && typePara.Definition.Name == "Type Name")
                                                    {
                                                        typeEle.Name = linkedTypeEle.Name;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                diffStr = typePara.Definition.Name + ",パラメーターを追加しました。";
                                                diff.S1 = typePara.Definition.Name;
                                                diff.S2 = "パラメーターを追加しました。";
                                            }
                                            break;
                                        case StorageType.ElementId:
                                            var valueEid = typePara.AsElementId();
                                            if (linkedTypePara != null)
                                            {
                                                ElementId linkedValueEid = linkedTypePara.AsElementId();
                                                if (valueEid != linkedValueEid)
                                                {
                                                    diffStr = typePara.Definition.Name + "," + linkedValueEid + "," + valueEid;
                                                    diff.S1 = typePara.Definition.Name;
                                                    diff.S2 = linkedValueEid.ToString();
                                                    diff.S3 = valueEid.ToString();

                                                    if (update && !typePara.IsReadOnly)
                                                    {
                                                        if (unchangedPars is null || (unchangedPars != null && !unchangedPars.Contains(typePara.Definition.Name)))
                                                        {
                                                            typeEle.LookupParameter(typePara.Definition.Name).Set(linkedValueEid);
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                diffStr = typePara.Definition.Name + ",パラメーターを追加しました。";
                                                diff.S1 = typePara.Definition.Name;
                                                diff.S2 = "パラメーターを追加しました。";
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                }
                                if (diffStr != "")
                                {
                                    tempDiffStrs.Add(diffStr);
                                    tempDiffs.Add(diff);
                                }
                            }
                            #endregion Type Parameters Checking

                        }
                        if (tempDiffStrs.Count > 0)
                        {
                            string eleDefStr = "TYPE," + typeEle.Category.Name + "," + typeEle.Name + "," + typeEle.Id;
                            Diff eleDef = new Diff("TYPE", typeEle.Category.Name, typeEle.Name, typeEle.Id.ToString());
                            diffStrs.Add(eleDefStr);
                            diffs.Add(eleDef);
                            tempDiffStrs.ForEach(s => diffStrs.Add(s));
                            tempDiffs.ForEach(d => diffs.Add(d));
                            var insts = new FilteredElementCollector(doc, doc.ActiveView.Id)
                                            .WhereElementIsNotElementType()
                                            .Where(e => e.GetTypeId().IntegerValue == typeEleId.IntegerValue)
                                            .ToList();
                            insts.ForEach(e => modifiedEles.Add(e));
                        }
                    }
                }
            }

            diffs.ForEach(d => col1.Add(d.S1));
            diffs.ForEach(d => col2.Add(d.S2));
            diffs.ForEach(d => col3.Add(d.S3));
            diffs.ForEach(d => col4.Add(d.S4));
            List<List<string>> cols = new List<List<string>> { col1, col2, col3, col4 };

            return (newEles, modifiedEles, deletedEles, diffStrs, diffs, cols);
        }
        public static List<CatFamPar> GetCatFamPars(Document doc, List<Element> eles)
        {
            if (eles is null || eles.Count == 0)
            {
                return null;
            }
            List<CatFamPar> catFamPars = new List<CatFamPar>();
            var eleCatGroups = eles.Where(e => !e.Name.Contains("増し打ち")
                                            && !e.Name.Contains("埋め戻し")
                                            && !e.Name.Contains("砕石"))
                                   .GroupBy(e => e.Category.Name);
            foreach (var eleCatGroup in eleCatGroups)
            {
                var eleFamGroups = eleCatGroup
                                    .Where(e => doc.GetElement(e.GetTypeId()) != null)
                                    .GroupBy(e => (doc.GetElement(e.GetTypeId()) as ElementType).FamilyName);
                foreach (var eleFamGroup in eleFamGroups)
                {
                    var type = new FilteredElementCollector(doc, doc.ActiveView.Id)
                                    .WhereElementIsNotElementType()
                                    .Where(e => doc.GetElement(e.GetTypeId()) != null && (doc.GetElement(e.GetTypeId()) as ElementType).FamilyName == eleFamGroup.Key)
                                    .First();
                    var inst = new FilteredElementCollector(doc, doc.ActiveView.Id)
                                    .WhereElementIsNotElementType()
                                    .Where(e => doc.GetElement(e.GetTypeId()) != null && (doc.GetElement(e.GetTypeId()) as ElementType).Name == type.Name)
                                    .First();
                    var typeParams = type.Parameters;
                    var instParams = inst.Parameters;
                    foreach (Parameter instParam in instParams)
                    {
                        if (!instParam.IsReadOnly && !instParam.Definition.Name.Contains("状態"))
                        {
                            CatFamPar catFamPar = new CatFamPar(eleCatGroup.Key, eleFamGroup.Key, "Instance Parameters", instParam.Definition.Name);
                            catFamPars.Add(catFamPar);
                        }
                    }
                    foreach (Parameter typeParam in typeParams)
                    {
                        if (!typeParam.IsReadOnly)
                        {
                            CatFamPar catFamPar = new CatFamPar(eleCatGroup.Key, eleFamGroup.Key, "Type Parameters", typeParam.Definition.Name);
                            catFamPars.Add(catFamPar);
                        }
                    }
                }
            }
            return catFamPars;
        }
        public static List<CatFamPar> SetDefaultPars(List<CatFamPar> catFamPars)
        {
            if (catFamPars is null || catFamPars.Count == 0)
            {
                return null;
            }
            List<CatFamPar> defaultList = new List<CatFamPar>();
            List<string> defPars = new List<string>()
            {
                "基礎_入隅",
                "基礎_内部",
                "基礎_出隅",
                "基礎_外周",
                "基礎_外部",
                "柱型_マーク",
                "柱型_合計",
                "埋め戻し",
                "柱型の上増し打ちの無視",
                "砕石なし",
                "フレーム_位置",
                "フレーム_カット長",
                "床_幅",
                "床_長",
                "床_ピット_マイナス_幅",
                "床_ピット_マイナス_長",
                "床_ピット_掘削"
            };
            foreach (var catFamPar in catFamPars)
            {
                if (defPars.Contains(catFamPar.Parameter))
                {
                    catFamPar.Unchange = true;
                }
                defaultList.Add(catFamPar);
            }
            return defaultList;
        }
        public class CopyUseDestination : IDuplicateTypeNamesHandler
        {
            public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
            {
                return DuplicateTypeAction.UseDestinationTypes;
            }
        }
    }
}
