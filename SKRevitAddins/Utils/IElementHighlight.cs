using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.Utils
{
    public static class FilterElementHelper
    {
        public static List<Element> Filter(UIDocument uidoc, ExcelRangeNameParseHelper.RangeObject rngObj)
        {
            Document doc = uidoc.Document;
            if (rngObj.ObjectType == ExcelRangeNameParseHelper.ObjectType.Reset)
            {
                ElementHighlightUtils.ResetHighlightElements(uidoc);
            }

            switch (rngObj.ObjectType)
            {
                //case ObjectType.IchijiKairyou:
                //    return FilterIchijiKairyou(doc, rngObj as RngObjIchijiKairyou);


                default:
                    break;
            }

            return null;
        }

        //private static List<Element> FilterIchijiKairyou(Document doc, RngObjIchijiKairyou rngObjIchijiKairyou)
        //{
        //    return new FilteredElementCollector(doc)
        //        .WhereElementIsNotElementType()
        //        .OfCategory(BuiltInCategory.OST_CurtainWallPanels)
        //        .Where(e => (Math.Round((e as FamilyInstance).Symbol.LookupParameter(TnfParameters.CurtainIchijiJibanDepth).AsDouble() * 304.8, 0).ToString().Equals(rngObjIchijiKairyou.Thickness))
        //        && !((e as FamilyInstance).Host.Name.Equals(TnfParameters.GaikoDes)))?.ToList();
        //}
        public static class ElementHighlightUtils
        {
            public static void HighlightElements(UIDocument uidoc, List<Element> eleList, byte r = 0, byte g = 255,
                byte b = 0)
            {
                //uidoc.Document.ActiveView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                List<ElementId> eleIdList = new List<ElementId>();
                OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                Color red = new Color(255, 0, 0);
                Element solidFill = new FilteredElementCollector(uidoc.Document).OfClass(typeof(FillPatternElement))
                    .Where(q => q.Name.Contains("Solid")).First();

                ogs.SetProjectionLineColor(red);
                ogs.SetProjectionLineWeight(8);
                ogs.SetSurfaceBackgroundPatternId(solidFill.Id);
                ogs.SetSurfaceBackgroundPatternColor(new Color(r, g, b));

                if ((eleList != null) && (eleList.Count > 0))
                {
                    eleList[0].Category.set_Visible(uidoc.ActiveView, true);
                    foreach (var ele in eleList)
                    {
                        eleIdList.Add(ele.Id);
                        uidoc.ActiveView.SetElementOverrides(ele.Id, ogs);
                    }
                }

                uidoc.RefreshActiveView();
                uidoc.Selection.SetElementIds(eleIdList);
                //uidoc.Document.ActiveView.IsolateElementsTemporary(eleIdList);
                uidoc.ShowElements(eleIdList);
            }

            public static void ResetHighlightElements(UIDocument uidoc)
            {
                var eleList = new FilteredElementCollector(uidoc.Document)
                    .WhereElementIsNotElementType().ToList();

                OverrideGraphicSettings ogs = new OverrideGraphicSettings();

                if ((eleList != null) && (eleList.Count > 0))
                {
                    foreach (var ele in eleList)
                    {
                        uidoc.ActiveView.SetElementOverrides(ele.Id, ogs);
                    }
                }

                uidoc.RefreshActiveView();
                uidoc.Document.ActiveView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
            }
        }
    }
}
