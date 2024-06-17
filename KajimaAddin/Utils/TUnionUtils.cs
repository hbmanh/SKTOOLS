using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKToolsAddins.Utils
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
