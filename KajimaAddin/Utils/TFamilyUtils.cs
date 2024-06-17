using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKToolsAddins.Utils
{
    class TFamilyUtils
    {
        public static FamilySymbol FindFamilySymbolByName(Document doc, string familyName)
        {
            var familySymbols = new FilteredElementCollector(doc).WherePasses(new ElementClassFilter(typeof(FamilySymbol))).Where(e => e.Name.Equals(familyName));
            if (familySymbols == null || familySymbols.Count() == 0)
            {
                return null;
            }
            else
            {
                return familySymbols.ElementAt(0) as FamilySymbol;
            }

        }
    }
}
