using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
{
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
                if (geoObject is GeometryInstance instance)
                {
                    foreach (GeometryObject instObj in instance.SymbolGeometry)
                    {
                        if (instObj is GeometryInstance blockInstance)
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

            return blocks;
        }
    }
}
