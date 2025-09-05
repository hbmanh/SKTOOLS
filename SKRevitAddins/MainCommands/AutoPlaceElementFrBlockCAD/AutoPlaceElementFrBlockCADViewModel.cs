//using System;
//using System.Collections.Generic;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.UI;

//namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
//{
//    public class AutoPlaceElementFrBlockCADViewModel
//    {
//        private readonly Document _doc;

//        public AutoPlaceElementFrBlockCADViewModel(Document doc)
//        {
//            _doc = doc;
//        }

//        public void ProcessCadLink(ImportInstance cadLink)
//        {
//            var options = new Options
//            {
//                ComputeReferences = false,
//                IncludeNonVisibleObjects = false,
//                DetailLevel = ViewDetailLevel.Coarse
//            };

//            GeometryElement geoElement = cadLink.get_Geometry(options);
//            if (geoElement == null) return;

//            foreach (GeometryObject geoObject in geoElement)
//            {
//                if (geoObject is GeometryInstance dwgInstance)
//                {
//                    foreach (GeometryObject instObj in dwgInstance.SymbolGeometry)
//                    {
//                        if (instObj is GeometryInstance blockInstance)
//                        {
//                            string symbolName = GetGeometryInstanceName(_doc, blockInstance);

//                            if (!string.Equals(symbolName, "dwg", StringComparison.InvariantCultureIgnoreCase) &&
//                                !string.IsNullOrWhiteSpace(symbolName))
//                            {
//                                // Ví dụ: hiển thị thông báo
//                                TaskDialog.Show("Block Found", $"Tên block: {symbolName}");
//                            }
//                        }
//                    }
//                }
//            }
//        }

//        private static string GetGeometryInstanceName(Document doc, GeometryInstance instance)
//        {
//            if (instance == null) return string.Empty;

//#if REVIT2024 || REVIT2025
//            SymbolGeometryId symbolGeoId = instance.GetSymbolGeometryId();
//            ElementId symbolId = symbolGeoId.SymbolId;
//            Element symbolElem = doc.GetElement(symbolId);
//            return symbolElem?.Name ?? string.Empty;
//#else
//            try
//            {
//                dynamic dynInstance = instance;
//                var symbol = dynInstance.Symbol;
//                return symbol?.Name ?? string.Empty;
//            }
//            catch
//            {
//                try
//                {
//                    ElementId symbolId = instance.SymbolId;
//                    Element symbolElem = doc.GetElement(symbolId);
//                    return symbolElem?.Name ?? string.Empty;
//                }
//                catch
//                {
//                    return string.Empty;
//                }
//            }
//#endif
//        }
//    }
//}
