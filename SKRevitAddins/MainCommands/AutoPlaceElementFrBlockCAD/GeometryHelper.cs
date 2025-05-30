﻿using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
{
    public class BlockWithLink
    {
        public GeometryInstance Block { get; set; }
        public ImportInstance CadLink { get; set; }
    }

    public class GeometryHelper
    {
        public static List<BlockWithLink> GetBlockNamesFromCadLink(ImportInstance cadLink)
        {
            var blocks = new List<BlockWithLink>();
            var options = new Options
            {
                ComputeReferences = false,
                IncludeNonVisibleObjects = false,
                DetailLevel = ViewDetailLevel.Coarse
            };

            GeometryElement geoElement = cadLink.get_Geometry(options);

            foreach (GeometryObject geoObject in geoElement)
            {
                if (geoObject is GeometryInstance dwgInstance)
                {
                    foreach (GeometryObject instObj in dwgInstance.SymbolGeometry)
                    {
                        if (instObj is GeometryInstance blockInstance)
                        {
                            string symbolName = blockInstance.Symbol?.Name ?? "";
                            if (!string.Equals(symbolName, "dwg", System.StringComparison.InvariantCultureIgnoreCase) &&
                                !string.IsNullOrWhiteSpace(symbolName))
                            {
                                blocks.Add(new BlockWithLink
                                {
                                    Block = blockInstance,
                                    CadLink = cadLink
                                });
                            }
                        }
                    }
                }
            }
            return blocks;
        }
    }
}
