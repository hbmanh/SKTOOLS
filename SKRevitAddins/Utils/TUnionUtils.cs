using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace SKRevitAddins.Utils
{
    class TUnionUtils
    {
        public static Solid UnionSolids(List<Solid> solids)
        {
            Solid unionSolid  = null;

            bool beginFlag = true;
            foreach (var solid in solids)
            {
                if (beginFlag)
                {
                    unionSolid = solid;
                    beginFlag = false;
                }
                else
                {
                    unionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(unionSolid, solid, BooleanOperationsType.Union);
                }

            }
            return unionSolid;
        }
    }
}
