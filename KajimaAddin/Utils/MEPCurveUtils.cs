using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKToolsAddins.Utils
{
    public static class MEPCurveUtils
    {
        public static int SelectPipeSize (this int fu, Dictionary<int, int> sizeDict)
        {
            int i = 0;
            var ele = sizeDict.ElementAt(i);
            var key = ele.Key;
            var value = ele.Value;
            while (fu > key)
            {
                i++;
                ele = sizeDict.ElementAt(i);
                key = ele.Key;
                value = ele.Value;
            }
            return value;
        }

        public static List<XYZ> FindIntersectionPoints(List<MEPCurve> mepCurves)
        {
            List<XYZ> xPoints = new List<XYZ>();
            foreach (var curve1 in mepCurves)
            {
                foreach (var curve2 in mepCurves)
                {
                    if (curve1.Id == curve2.Id) continue;

                    if (!IsIntersection(curve1, curve2, out XYZ intersectionPoint)) continue;
                    // Kiểm tra nếu intersectionPoint chưa tồn tại trong danh sách XPoints
                    if (!xPoints.Any(p => p.IsAlmostEqualTo(intersectionPoint)))
                    {
                        // Add intersection points
                        xPoints.Add(intersectionPoint);
                    }
                }
            }
            return xPoints;
        }
        public static bool IsIntersection(MEPCurve element1, MEPCurve element2, out XYZ intersectionPoint)
        {
            intersectionPoint = null;

            var curve1 = (element1.Location as LocationCurve).Curve;
            var curve2 = (element2.Location as LocationCurve).Curve;

            IntersectionResultArray results;
            SetComparisonResult result = curve1.Intersect(curve2, out results);

            if (result == SetComparisonResult.Overlap && results != null && results.Size > 0)
            {
                intersectionPoint = results.get_Item(0).XYZPoint;
                return true;
            }

            return false;
        }
        public static List<(MEPCurve, MEPCurve)> GetCollinearCurves(List<MEPCurve> mepCurves)
        {
            var collinearCurves = new List<(MEPCurve, MEPCurve)>();

            for (int i = 0; i < mepCurves.Count; i++)
            {
                for (int j = i + 1; j < mepCurves.Count; j++)
                {
                    var mepCurveI = mepCurves[i];
                    var mepCurveJ = mepCurves[j];

                    if (AreCurvesCollinear(mepCurveI, mepCurveJ))
                    {
                        collinearCurves.Add((mepCurveI, mepCurveJ));
                    }
                }
            }

            return collinearCurves;
        }

        // Kiểm tra xem hai MEPCurve có thẳng hàng không
        public static bool AreCurvesCollinear(MEPCurve curve1, MEPCurve curve2)
        {
            Line line1 = (curve1.Location as LocationCurve)?.Curve as Line;
            Line line2 = (curve2.Location as LocationCurve)?.Curve as Line;

            if (line1 == null || line2 == null)
                return false;

            XYZ dir1 = (line1.GetEndPoint(1) - line1.GetEndPoint(0)).Normalize();
            XYZ dir2 = (line2.GetEndPoint(1) - line2.GetEndPoint(0)).Normalize();

            // Kiểm tra tích vô hướng, nếu gần bằng ±1 thì hai vector song song (thẳng hàng)
            return Math.Abs(dir1.DotProduct(dir2)) > 0.99;
        }

        public static MEPCurve CreateNewCurveFromCurvesCollinear(Document doc, MEPCurve curve1, MEPCurve curve2, Level level)
        {
            if (!IsElementValid(curve1) || !IsElementValid(curve2))
            {
                return null;
            }

            var line1 = (curve1.Location as LocationCurve)?.Curve as Line;
            var line2 = (curve2.Location as LocationCurve)?.Curve as Line;

            if (line1 == null || line2 == null)
            {
                return null;
            }

            XYZ start1 = line1.GetEndPoint(0);
            XYZ end1 = line1.GetEndPoint(1);
            XYZ start2 = line2.GetEndPoint(0);
            XYZ end2 = line2.GetEndPoint(1);

            // Find the two farthest points
            XYZ[] points = { start1, end1, start2, end2 };
            double maxDistance = 0;
            XYZ point1 = null;
            XYZ point2 = null;

            for (int p1 = 0; p1 < points.Length; p1++)
            {
                for (int p2 = p1 + 1; p2 < points.Length; p2++)
                {
                    double distance = points[p1].DistanceTo(points[p2]);
                    if (distance > maxDistance)
                    {
                        maxDistance = distance;
                        point1 = points[p1];
                        point2 = points[p2];
                    }
                }
            }

            // Create a new pipe or duct from the two farthest points
            MEPCurve newCurve = null;
            if (curve1 is Pipe)
            {
                newCurve = Pipe.Create(doc, curve1.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId(), curve1.GetTypeId(), level.Id, point1, point2);
                newCurve.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(curve1.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble());
            }
            else if (curve1 is Duct)
            {
                newCurve = Duct.Create(doc, curve1.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId(), curve1.GetTypeId(), level.Id, point1, point2);
                newCurve.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).Set(curve1.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).AsDouble());
            }

            return newCurve;
        }

        public static bool IsElementValid(Element element)
        {
            try
            {
                var id = element.Id;
                var doc = element.Document;
                return true;
            }
            catch
            {
                return false;
            }
        }



        public static List<MEPCurve> SplitCurve(Document doc, MEPCurve curve, List<XYZ> xPoints, Level level)
        {
            List<MEPCurve> splitCurves = new List<MEPCurve>();

            // Lấy curve gốc
            var mainCurve = (curve.Location as LocationCurve)?.Curve;
            if (mainCurve == null)
                return splitCurves;

            var mainStart = mainCurve.GetEndPoint(0);
            var mainEnd = mainCurve.GetEndPoint(1);
            var typeId = curve.GetTypeId();
            var systemId = curve is Pipe
                            ? curve.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
                            : curve.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();
   
            // Lọc các điểm giao nằm trên curve hiện tại
            List<XYZ> validXPoints = xPoints.Where(p => IsPointOnCurve(mainCurve, p)).ToList();

            // Thêm các điểm đầu và cuối của curve vào danh sách points
            List<XYZ> points = new List<XYZ> { mainStart };
            points.AddRange(validXPoints.OrderBy(p => p.DistanceTo(mainStart)).ToList());
            points.Add(mainEnd);

            // Tạo các đoạn curve mới tại các điểm intersection
            for (int i = 0; i < points.Count - 1; i++)
            {
                XYZ startPoint = points[i];
                XYZ endPoint = points[i + 1];
                if (startPoint.DistanceTo(endPoint) >= 0.1) // Ensure the length is at least 1/10 inch
                {
                    MEPCurve newCurve;
                    if (curve is Pipe)
                    {
                        newCurve = Pipe.Create(doc, systemId, typeId, level.Id, startPoint, endPoint);
                        newCurve.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(UnitUtils.MmToFeet(20));
                    }
                    else if (curve is Duct)
                    {
                        newCurve = Duct.Create(doc, systemId, typeId, level.Id, startPoint, endPoint);
                        newCurve.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).Set(UnitUtils.MmToFeet(75));
                    }
                    else
                    {
                        continue;
                    }
                    newCurve.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(UnitUtils.MmToFeet(2800));
                    splitCurves.Add(newCurve);
                }
            }

            // Xóa curve cũ
            doc.Delete(curve.Id);

            return splitCurves;
        }
        public static List<MEPCurve> SplitCPlaceholders(Document doc, MEPCurve curve, List<XYZ> xPoints, Level level)
        {
            List<MEPCurve> splitCurves = new List<MEPCurve>();

            // Lấy curve gốc
            var mainCurve = (curve.Location as LocationCurve)?.Curve;
            if (mainCurve == null)
                return splitCurves;

            var mainStart = mainCurve.GetEndPoint(0);
            var mainEnd = mainCurve.GetEndPoint(1);
            var typeId = curve.GetTypeId();
            var systemId = curve is Pipe
                            ? curve.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
                            : curve.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();

            // Lọc các điểm giao nằm trên curve hiện tại
            List<XYZ> validXPoints = xPoints.Where(p => IsPointOnCurve(mainCurve, p)).ToList();

            // Thêm các điểm đầu và cuối của curve vào danh sách points
            List<XYZ> points = new List<XYZ> { mainStart };
            points.AddRange(validXPoints.OrderBy(p => p.DistanceTo(mainStart)).ToList());
            points.Add(mainEnd);

            // Tạo các đoạn curve mới tại các điểm intersection
            for (int i = 0; i < points.Count - 1; i++)
            {
                XYZ startPoint = points[i];
                XYZ endPoint = points[i + 1];
                if (startPoint.DistanceTo(endPoint) >= 0.1) // Ensure the length is at least 1/10 inch
                {
                    MEPCurve newCurve;
                    if (curve is Pipe)
                    {
                        newCurve = Pipe.CreatePlaceholder(doc, systemId, typeId, level.Id, startPoint, endPoint);
                        newCurve.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(UnitUtils.MmToFeet(20));
                    }
                    else if (curve is Duct)
                    {
                        newCurve = Duct.CreatePlaceholder(doc, systemId, typeId, level.Id, startPoint, endPoint);
                        newCurve.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).Set(UnitUtils.MmToFeet(75));
                    }
                    else
                    {
                        continue;
                    }
                    newCurve.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(UnitUtils.MmToFeet(2800));
                    splitCurves.Add(newCurve);
                }
            }

            // Xóa curve cũ
            doc.Delete(curve.Id);

            return splitCurves;
        }
        public static bool IsPointOnCurve(Curve curve, XYZ point)
        {
            var projectedPoint = curve.Project(point).XYZPoint;
            return point.IsAlmostEqualTo(projectedPoint);
        }
    }
}
