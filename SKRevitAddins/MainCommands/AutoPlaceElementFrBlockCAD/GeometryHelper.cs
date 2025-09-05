//using System;
//using System.Collections.Generic;
//using Autodesk.Revit.DB;

//namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
//{
//    public class BlockWithLink
//    {
//        public GeometryInstance Block { get; set; }
//        public ImportInstance CadLink { get; set; }
//    }

//    public class GeometryHelper
//    {
//        public static List<BlockWithLink> GetBlockNamesFromCadLink(ImportInstance cadLink)
//        {
//            var blocks = new List<BlockWithLink>();

//            var options = new Options
//            {
//                ComputeReferences = false,
//                IncludeNonVisibleObjects = false,
//                DetailLevel = ViewDetailLevel.Coarse
//            };

//            GeometryElement geoElement = cadLink.get_Geometry(options);
//            if (geoElement == null) return blocks;

//            foreach (GeometryObject geoObject in geoElement)
//            {
//                if (geoObject is GeometryInstance dwgInstance)
//                {
//                    foreach (GeometryObject instObj in dwgInstance.SymbolGeometry)
//                    {
//                        if (instObj is GeometryInstance blockInstance)
//                        {
//                            string symbolName = GetGeometryInstanceName(cadLink.Document, blockInstance);

//                            if (!string.Equals(symbolName, "dwg", StringComparison.InvariantCultureIgnoreCase) &&
//                                !string.IsNullOrWhiteSpace(symbolName))
//                            {
//                                blocks.Add(new BlockWithLink
//                                {
//                                    Block = blockInstance,
//                                    CadLink = cadLink
//                                });
//                            }
//                        }
//                    }
//                }
//            }

//            return blocks;
//        }

//        /// <summary>
//        /// Lấy tên Symbol từ GeometryInstance, tương thích nhiều phiên bản Revit
//        /// </summary>
//        private static string GetGeometryInstanceName(Document doc, GeometryInstance instance)
//        {
//            if (instance == null) return string.Empty;

//#if REVIT2024 || REVIT2025
//    // API mới
//    SymbolGeometryId symbolGeoId = instance.GetSymbolGeometryId();
//    ElementId symbolId = symbolGeoId.SymbolId;
//    Element symbolElem = doc.GetElement(symbolId);
//    return symbolElem?.Name ?? string.Empty;
//#else
//            // API cũ
//            ElementId symbolId = instance.SymbolId;
//            Element symbolElem = doc.GetElement(symbolId);
//            return symbolElem?.Name ?? string.Empty;
//#endif
//        }

//    }
//}
