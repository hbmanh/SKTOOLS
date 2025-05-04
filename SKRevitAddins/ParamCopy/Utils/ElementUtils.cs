using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParamCopy
{
    public static class ElementUtils
    {
        public static Element GetElementType(this Element instance, Document doc)
        {
            if (instance == null || doc == null) return null;
            var instanceTypeId = instance.GetTypeId();
            return doc.GetElement(instanceTypeId);
        }
    }
}
