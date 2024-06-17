using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace SKToolsAddins.Utils
{
    public static class PointUtils
    {
        public static bool CheckCoincidentPoints(this XYZ p1, XYZ p2, bool ignoreZ = false)
        {
            bool isCoincident = false;
            if (Math.Round(p1.DistanceTo(p2), 2) == 0)
            {
                isCoincident = true;
                return isCoincident;
            }
            if (ignoreZ)
            {
                XYZ vec = p2 - p1;
                double angle = Math.Round(vec.AngleTo(XYZ.BasisZ), 3);
                if (angle == 0 || Math.Round(angle - Math.PI, 3) == 0)
                {
                    isCoincident = true;
                }
            }
            return isCoincident;
        }
        public static bool CheckColinearPoints(this XYZ p1, XYZ p2, XYZ p3)
        {
            if ((p1.CheckCoincidentPoints(p2) is false)
                && (p2.CheckCoincidentPoints(p3) is false)
                && (p3.CheckCoincidentPoints(p1) is false))
            {
                double val = (p2.Y - p1.Y) * (p3.X - p2.X) - (p2.X - p1.X) * (p3.Y - p2.Y);
                if (Math.Round(val, 3) == 0)
                {
                    return true;
                }
                else return false;
            }
            else return false;
        }
        public static XYZ GetMidColinearPoints(this XYZ p1, XYZ p2, XYZ p3)
        {
            XYZ midPoint = p1;
            bool isColinear = CheckColinearPoints(p1, p2, p3);
            if (isColinear is true)
            {
                var distp1 = p1.DistanceTo(p2) + p1.DistanceTo(p3);
                var distp2 = p2.DistanceTo(p1) + p2.DistanceTo(p3);
                var distp3 = p3.DistanceTo(p1) + p3.DistanceTo(p2);
                var minDist = distp1;
                if (minDist > distp2)
                {
                    minDist = distp2;
                    midPoint = p2;
                }
                if (minDist > distp3)
                {
                    minDist = distp3;
                    midPoint = p3;
                }
                return midPoint;
            }
            else return null;
            
        }
        public static bool IsPointInLine(this XYZ point, Line line)
        {
            if (point is null || line is null)
            {
                return false;
            }
            var p1 = line.GetEndPoint(0);
            var p2 = line.GetEndPoint(1);
            bool colinear = point.CheckColinearPoints(p1, p2);
            if (!colinear) return false;
            XYZ mid = point.GetMidColinearPoints(p1, p2);
            bool isMid = mid.CheckCoincidentPoints(point);
            if (isMid) return true; else return false;
        }
        public static List<XYZ> NoColinearList(List<XYZ> pointList)
        {
            List<XYZ> noColinearList = new List<XYZ>();
            foreach (var p in pointList)
            {
                noColinearList.Add(p);
            }
            foreach (var p1 in pointList)
            {
                foreach (var p2 in pointList)
                {
                    foreach (var p3 in pointList)
                    {
                        if ((p1 != p2) && (p2 != p3) && (p3 != p1))
                        {
                            bool isColinear = CheckColinearPoints(p1, p2, p3);
                            if (isColinear is true)
                            {
                                XYZ midPoint = GetMidColinearPoints(p1, p2, p3);
                                if (noColinearList.Contains(midPoint))
                                {
                                    noColinearList.Remove(midPoint);
                                }
                            }
                        }
                    }
                }
            }
            return noColinearList;
        }
        public static CurveLoop GetSolidCurveLoop(this Solid solid)
        {
            var floorTopFace = solid.GetTopPlanarFace();
            return floorTopFace.GetEdgesAsCurveLoops().First();
        }
        public static CurveLoop GetFloorCurveLoop (this Element floor, Document doc)
        {
            var floorSolid = floor.GetAllSolidsAdvance(true)
                .Where(s => (s != null) && (s.Volume > 0))
                .First();
            List<Element> foundColl = new FilteredElementCollector(doc).WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .OfClass(typeof(FamilyInstance)).Cast<Element>()
                .Where(e => (e as FamilyInstance).Symbol.FamilyName.Contains("基礎")).ToList();
            var foundCollSolidUnion = foundColl.UnionInstColl();
            if ((foundCollSolidUnion != null) && (foundCollSolidUnion.Volume > 0))
            {
                floorSolid = BooleanOperationsUtils.ExecuteBooleanOperation(floorSolid, foundCollSolidUnion, BooleanOperationsType.Difference);
            }
            return floorSolid.GetSolidCurveLoop();
        }
        public static List<XYZ> GetSolidPoints(this Solid solid)
        {
            List<XYZ> solidPoints = new List<XYZ>();

            var solidCurveLoop = solid.GetSolidCurveLoop();

            foreach (Curve floorCurve in solidCurveLoop)
            {
                var p1 = floorCurve.GetEndPoint(0);
                var p2 = floorCurve.GetEndPoint(1);

                bool isP1Included = false;
                bool isP2Included = false;

                foreach (var floorPoint in solidPoints)
                {
                    if (Math.Round(p1.DistanceTo(floorPoint), 3) == 0)
                    {
                        isP1Included = true;
                    }
                    if (Math.Round(p1.DistanceTo(floorPoint), 3) == 0)
                    {
                        isP2Included = true;
                    }
                }
                if (!isP1Included)
                {
                    solidPoints.Add(p1);
                }
                if (!isP2Included)
                {
                    solidPoints.Add(p2);
                }
            }
            return solidPoints;
        }
        public static List<XYZ> GetFloorPoints(this Element floor, Document doc)
        {
            List<XYZ> floorPoints = new List<XYZ>();

            var floorCurveLoop = floor.GetFloorCurveLoop(doc);
            
            foreach (Curve floorCurve in floorCurveLoop)
            {
                var p1 = floorCurve.GetEndPoint(0);
                var p2 = floorCurve.GetEndPoint(1);

                bool isP1Included = false;
                bool isP2Included = false;

                foreach (var floorPoint in floorPoints)
                {
                    if (Math.Round(p1.DistanceTo(floorPoint), 3) == 0)
                    {
                        isP1Included = true;
                    }
                    if (Math.Round(p1.DistanceTo(floorPoint), 3) == 0)
                    {
                        isP2Included = true;
                    }
                }
                if (!isP1Included)
                {
                    floorPoints.Add(p1);
                }
                if (!isP2Included)
                {
                    floorPoints.Add(p2);
                }
            }
            return floorPoints;
        }
        public static bool CheckTwoCurvesIfTheyAreSame(this Curve curve1, Curve curve2)
        {
            var p11 = curve1.GetEndPoint(0);
            var p12 = curve1.GetEndPoint(1);
            var p21 = curve2.GetEndPoint(0);
            var p22 = curve2.GetEndPoint(1);
            if (((Math.Round(p11.DistanceTo(p21), 3) == 0) && (Math.Round(p12.DistanceTo(p22), 3) == 0))
                                || ((Math.Round(p11.DistanceTo(p22), 3) == 0) && (Math.Round(p12.DistanceTo(p21), 3) == 0)))
            {
                return true;
            }
            else return false;
        }
        public static (bool, XYZ) CheckTwoCurveIfSameOnePoint(this Curve curve1, Curve curve2)
        {
            if (curve1.CheckTwoCurvesIfTheyAreSame(curve2))
            {
                return (false, null);
            }
            var p11 = curve1.GetEndPoint(0);
            var p12 = curve1.GetEndPoint(1);
            var p21 = curve2.GetEndPoint(0);
            var p22 = curve2.GetEndPoint(1);
            if (Math.Round(p11.DistanceTo(p21), 3) == 0)
            {
                return (true, p21);
            }
            else if (Math.Round(p11.DistanceTo(p22), 3) == 0)
            {
                return (true, p22);
            }
            else if (Math.Round(p12.DistanceTo(p21), 3) == 0)
            {
                return (true, p21);
            }
            else if (Math.Round(p12.DistanceTo(p22), 3) == 0)
            {
                return (true, p22);
            }
            else return (false, null);
        }
        public static XYZ ShiftCoor(this XYZ point, XYZ newCoor)
        {
            return new XYZ(point.X - newCoor.X, point.Y - newCoor.Y, point.Z - newCoor.Z);
        }

    }
}
