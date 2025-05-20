using System;
using Autodesk.Revit.DB;

namespace SKRevitAddins.PermissibleRangeFrame
{
    internal static class MEPCurveExtensions
    {
        /*──────────  Đơn vị  ──────────*/
        public static double ToInternalUnits(this double mm)
        {
            // dùng helper có sẵn trong SKRevitAddins.Utils
            return SKRevitAddins.Utils.UnitUtils.MmToFeet(mm);
        }

        /*──────────  Kích thước ống / duct  ──────────*/
        public static double GetOuterDiameter(this MEPCurve mep)
        {
            if (mep == null) return 0;
            double P(string n) => mep.LookupParameter(n)?.AsDouble() ?? 0;

            double d = P("Diameter");                       // pipe tròn
            if (d > 0) return d;

            double w = P("Width"), h = P("Height");         // duct chữ nhật
            if (w > 0 && h > 0) return Math.Max(w, h);

            double r = P("Radius");                         // một số family duct tròn
            if (r > 0) return 2 * r;

            return 0;
        }

        /*──────────  Kiểm điểm nằm trong Solid  ──────────*/
        public static bool IsPointInside(this Solid solid, XYZ p, double tol = 1e-4)
        {
            if (solid == null || solid.Volume == 0) return false;

            var bb = solid.GetBoundingBox();
            if (!bb.Contains(p, tol)) return false;         // ngoài BBox → chắc chắn ngoài

            /*  Với DirectShape hình hộp mỏng (extrusion 10 mm) :
                chỉ cần nằm trong BBox coi như “bên trong”.                                  */
            return true;
        }

        /*──────────  BoundingBox helper  ──────────*/
        private static bool Contains(this BoundingBoxXYZ bb, XYZ p, double tol)
        {
            return p.X > bb.Min.X - tol && p.X < bb.Max.X + tol &&
                   p.Y > bb.Min.Y - tol && p.Y < bb.Max.Y + tol &&
                   p.Z > bb.Min.Z - tol && p.Z < bb.Max.Z + tol;
        }
    }
}
