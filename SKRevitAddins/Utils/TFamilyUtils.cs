using System.Linq;
using Autodesk.Revit.DB;

namespace SKRevitAddins.Utils
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
