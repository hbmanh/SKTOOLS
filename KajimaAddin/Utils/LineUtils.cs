using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKToolsAddins.Utils
{
    public static class LineUtils
    {
        public static (Curve, Curve) SplitLineByTwoLines(Curve line1, Curve line2, Curve line3)
        {
            var line1Sp = line1.GetEndPoint(0);
            var line1Ep = line1.GetEndPoint(1);
            XYZ intersect12Point = new XYZ();
            XYZ intersect13Point = new XYZ();
            double distS12 = 0.0;
            double distS13 = 0.0;

            IntersectionResultArray intersect12Results;
            SetComparisonResult intersect12 = line1.Intersect(line2, out intersect12Results);
            if (intersect12 is SetComparisonResult.Overlap)
            {
                intersect12Point = intersect12Results.get_Item(0).XYZPoint;
                distS12 = line1Sp.DistanceTo(intersect12Point);
            }

            IntersectionResultArray intersect13Results;
            SetComparisonResult intersect13 = line1.Intersect(line3, out intersect13Results);
            if (intersect13 is SetComparisonResult.Overlap)
            {
                intersect13Point = intersect13Results.get_Item(0).XYZPoint;
                distS13 = line1Sp.DistanceTo(intersect13Point);
            }

            XYZ pointForSp = new XYZ();
            XYZ pointForEp = new XYZ();
            XYZ originPoint = new XYZ(0, 0, 0);

            Curve curve11 = null;
            Curve curve12 = null;

            if (!(distS12.Equals(0)) && !(distS13.Equals(0)))
            {
                if (distS12 < distS13)
                {
                    pointForSp = intersect12Point;
                    pointForEp = intersect13Point;
                }
                else
                {
                    pointForSp = intersect13Point;
                    pointForEp = intersect12Point;
                }
            }

            if (!(pointForSp is null) && (pointForSp.DistanceTo(originPoint) > 0) && !(pointForEp is null) && (pointForEp.DistanceTo(originPoint) > 0))
            {
                try
                {
                    curve11 = Line.CreateBound(line1Sp, pointForSp) as Curve;
                    curve12 = Line.CreateBound(line1Ep, pointForEp) as Curve;
                }
                catch (Exception)
                {
                }

            }

            return (curve11, curve12);
        }
        public static (Line, Line) SplitLineByOneLine(this Line line1, Line line2)
        {
            if (line1 is null || line2 is null)
            {
                return (null, null);
            }
            var line1Sp = line1.GetEndPoint(0);
            var line1Ep = line1.GetEndPoint(1);
            XYZ xPoint = new XYZ();
            double dist1 = 0.0;
            double dist2 = 0.0;
            IntersectionResultArray xResultArr;
            SetComparisonResult intersect = line1.Intersect(line2, out xResultArr);
            if (intersect is SetComparisonResult.Overlap)
            {
                xPoint = xResultArr.get_Item(0).XYZPoint;
            }
            if (!xPoint.CheckCoincidentPoints(line1Sp) && !xPoint.CheckCoincidentPoints(line1Ep))
            {
                var split1 = Line.CreateBound(line1Sp, xPoint);
                var split2 = Line.CreateBound(line1Ep, xPoint);
                return (split1, split2);
            }
            return (null, null);
        }
        public static ModelCurveArray GetMaxProfile(ModelCurveArrArray jibanProfiles)
        {
            ModelCurveArray jibanMaxProfile = new ModelCurveArray();
            foreach (ModelCurveArray jibanProfile in jibanProfiles)
            {
                if (LengthOfModelCurveArr(jibanProfile) > LengthOfModelCurveArr(jibanMaxProfile))
                {
                    jibanMaxProfile = jibanProfile;
                }
            }
            return jibanMaxProfile;
        }
        public static double LengthOfModelCurveArr(ModelCurveArray mca)
        {
            double totalLength = 0.0;
            foreach (ModelCurve m in mca)
            {
                var mLength = m.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                totalLength += mLength;
            }
            return totalLength;
        }
        public static Face GetHostFaceOfLine(this Curve curve, FaceArray faceArr)
        {
            Face host = null;
            foreach (var face in faceArr)
            {
                if (face is CylindricalFace) continue;
                var p1 = curve.GetEndPoint(0);
                var p2 = curve.GetEndPoint(1);
                Plane plane = Plane.CreateByNormalAndOrigin((face as PlanarFace).FaceNormal, (face as PlanarFace).Origin);
                plane.Project(p1, out UV uv1, out double dist1);
                plane.Project(p2, out UV uv2, out double dist2);
                if ((Math.Round(dist1, 2) == 0) && (Math.Round(dist2, 2) == 0))
                    host = face as PlanarFace;
            }
            return host;
        }
      
        public static bool CheckSameTwoCurves(this CurveElement curve1, CurveElement curve2)
        {
            if ((curve1 == null)
                || (curve2 == null))
            {
                return false;
            }

            IntersectionResultArray point;
            var result = curve1.GeometryCurve.Intersect(curve2.GeometryCurve, out point);
            if (result == SetComparisonResult.Equal)
            {
                return true;
            }
            return false;
        }
        public static bool IsCurveInsideAnother(this Curve curve1, Curve curve2, bool ignoreZ = false)
        {
            if (curve1 == null || curve2 == null || !curve1.AreTwoCurvesAligned(curve2))
            {
                return false;
            }
            XYZ c1P1 = curve1.GetEndPoint(0);
            XYZ c1P2 = curve1.GetEndPoint(1);
            XYZ c2P1 = curve2.GetEndPoint(0);
            XYZ c2P2 = curve2.GetEndPoint(1);

            XYZ midP1 = c1P1.GetMidColinearPoints(c2P1, c2P2);
            bool isMidP1 = (c1P1.CheckCoincidentPoints(midP1, ignoreZ) || c1P1.CheckCoincidentPoints(c2P1, ignoreZ) || c1P1.CheckCoincidentPoints(c2P2, ignoreZ)) ? true : false;

            XYZ midP2 = c1P2.GetMidColinearPoints(c2P1, c2P2);
            bool isMidP2 = (c1P2.CheckCoincidentPoints(midP2, ignoreZ) || c1P2.CheckCoincidentPoints(c2P1, ignoreZ) || c1P2.CheckCoincidentPoints(c2P2, ignoreZ)) ? true : false;

            if (isMidP1 && isMidP2)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static bool AreTwoCurvesAligned(this Curve curve1, Curve curve2)
        {
            if (curve1 == null || curve2 == null)
            {
                return false;
            }
            XYZ curve1P1 = curve1.GetEndPoint(0);
            XYZ curve1P2 = curve1.GetEndPoint(1);
            XYZ curve2P1 = curve2.GetEndPoint(0);
            XYZ curve2P2 = curve2.GetEndPoint(1);
            bool cond1 = curve1P1.CheckColinearPoints(curve2P1, curve2P2);
            bool cond2 = curve1P2.CheckColinearPoints(curve2P1, curve2P2);
            if (cond1 && cond2)
            {
                return true;
            }
            return false;
        }
        public static bool IsCurveInList(this Curve checkingCurve, List<Curve> curveList)
        {
            if ((checkingCurve == null)
                || (curveList == null)
                || (curveList.Count == 0))
            {
                return false;
            }
            foreach (var curve in curveList)
            {
                IntersectionResultArray point;
                var result = curve.Intersect(checkingCurve, out point);
                if (result != SetComparisonResult.Equal)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
        }
        public static List<CurveElement> RemoveCurveInCurveList (this List<CurveElement> curveList, CurveElement removecurve)
        {
            if ((curveList == null) 
                || (curveList.Count == 0) 
                || (removecurve == null))
            {
                return null;
            }
            List<CurveElement> newCurveList = new List<CurveElement>();
            foreach (var curve in curveList)
            {
                IntersectionResultArray point;
                var result = curve.GeometryCurve.Intersect(removecurve.GeometryCurve, out point);
                if (result != SetComparisonResult.Equal)
                {
                    newCurveList.Add(curve);
                }
            }
            return newCurveList;
        }
        public static (bool, List<Curve>, List<Curve>) ResultsOfCurveXCurve(this Curve originalCurve, Curve curve)
        {
            bool isIntersect = false;
            List<Curve> resultOriginalCurveList = new List<Curve>();
            List<Curve> resultCurveList = new List<Curve>();
            IntersectionResultArray resultArr = new IntersectionResultArray();
            SetComparisonResult result = originalCurve.Intersect(curve, out resultArr);
            if (result == SetComparisonResult.Equal)
            {
                var p1 = originalCurve.GetEndPoint(0);
                var p2 = originalCurve.GetEndPoint(1);
                var p3 = curve.GetEndPoint(0);
                var p4 = curve.GetEndPoint(1);

                List<XYZ> points = new List<XYZ>();
                points.Add(p1);
                points.Add(p2);
                points.Add(p3);
                points.Add(p4);

                points = points.OrderBy(p => Math.Round(p.X, 2)).ThenBy(p => Math.Round(p.Y, 2)).ToList();
                var startDist = Math.Round(points.ElementAt(0).DistanceTo(points.ElementAt(1)), 2);
                var midDist = Math.Round(points.ElementAt(1).DistanceTo(points.ElementAt(2)), 2);
                var endDist = Math.Round(points.ElementAt(2).DistanceTo(points.ElementAt(3)), 2);

                if (midDist > 0)
                {
                    isIntersect = true;
                    if (startDist > 0)
                    {
                        var startLine = Line.CreateBound(points.ElementAt(0), points.ElementAt(1));
                        if ((originalCurve as Line).CheckOneLineBelongToAnother(startLine))
                        {
                            resultOriginalCurveList.Add(startLine);
                        }
                        else if ((curve as Line).CheckOneLineBelongToAnother(startLine))
                        {
                            resultCurveList.Add(startLine);
                        }
                    }
                    if (endDist > 0)
                    {
                        var endLine = Line.CreateBound(points.ElementAt(2), points.ElementAt(3));
                        if ((originalCurve as Line).CheckOneLineBelongToAnother(endLine))
                        {
                            resultOriginalCurveList.Add(endLine);
                        }
                        else if ((curve as Line).CheckOneLineBelongToAnother(endLine))
                        {
                            resultCurveList.Add(endLine);
                        }
                    }
                }
            }
            return (isIntersect, resultOriginalCurveList, resultCurveList);
        }
        public static List<Curve> ResultsOfCurveXCurveList(this Curve originalCurve, List<Curve> curveList)
        {
            List<Curve> xCurveList = new List<Curve>();
            xCurveList.Add(originalCurve);
            int i = 0;
            while (i < xCurveList.Count)
            {
                var curCurve = xCurveList.ElementAt(i);
                bool intersectOnce = false;
                foreach (var curve in curveList)
                {
                    var curResult = curCurve.ResultsOfCurveXCurve(curve);
                    bool isIntersect = curResult.Item1;
                    List<Curve> curResultCuveList = curResult.Item2;
                    if (isIntersect)
                    {
                        xCurveList.RemoveAt(i);
                        if (curResultCuveList.Count > 0)
                        {
                            foreach (var curResultCurve in curResultCuveList)
                            {
                                xCurveList.Add(curResultCurve);
                            }
                        }
                        i = 0;
                        intersectOnce = true;
                        break;
                    }
                }
                if (intersectOnce is false)
                {
                    i++;
                }
            }

            return xCurveList;
        }
        public static List<Curve> ResultsOfCurveListXCurveList(this List<Curve> originalCurveList, List<Curve> curveList)
        {
            List<Curve> xCurveList = new List<Curve>();
            foreach (var originalCurve in originalCurveList)
            {
                List<Curve> curXCurveList = originalCurve.ResultsOfCurveXCurveList(curveList);
                if ((curXCurveList != null) && (curXCurveList.Count > 0))
                {
                    foreach (var curXCurve in curXCurveList)
                    {
                        xCurveList.Add(curXCurve);
                    }
                }
            }
            return xCurveList;
        }
        public static Line ProjectLineToPlane(this Line line, Plane plane)
        {
            UV uv1, uv2 = new UV();
            plane.Project(line.GetEndPoint(0), out uv1, out double dist1);
            plane.Project(line.GetEndPoint(1), out uv2, out double dist2);
            XYZ xyz1 = plane.Origin + (uv1.U * plane.XVec) + (uv1.V * plane.YVec);
            XYZ xyz2 = plane.Origin + (uv2.U * plane.XVec) + (uv2.V * plane.YVec);

            Line projectedLine = Line.CreateBound(xyz1, xyz2);
            return projectedLine;
        }
        public static XYZ GetIntersectBwLineAndFace(this Line line, Face face)
        {
            UV uv1, uv2 = new UV();

            Plane plane = Plane.CreateByNormalAndOrigin((face as PlanarFace).FaceNormal, (face as PlanarFace).Origin);

            plane.Project(line.Origin, out uv1, out double d);
            plane.Project(line.Origin + line.Direction, out uv2, out double b);

            XYZ xyz1 = plane.Origin + (uv1.U * plane.XVec) + (uv1.V * plane.YVec);
            XYZ xyz2 = plane.Origin + (uv2.U * plane.XVec) + (uv2.V * plane.YVec);

            if (Math.Round(xyz1.DistanceTo(xyz2), 2) == 0)
            {
                return xyz1;
            }
            else
            {
                Line projectedLine = Line.CreateUnbound(xyz1, xyz2 - xyz1);
                IntersectionResultArray iResult = new IntersectionResultArray();
                if (line.Intersect(projectedLine, out iResult) != SetComparisonResult.Disjoint)
                    return iResult.get_Item(0).XYZPoint;
                else return null;
            }
        }
        public static bool CheckOneLineBelongToAnother(this Line parentLine, Line checkingLine)
        {
            if ((parentLine is null) || (checkingLine is null))
            {
                return false;
            }
            bool isChildLine = false;

            var p1 = parentLine.GetEndPoint(0);
            var p2 = parentLine.GetEndPoint(1);
            var q1 = checkingLine.GetEndPoint(0);
            var q2 = checkingLine.GetEndPoint(1);

            var midPoint1 = q1.GetMidColinearPoints(p1, p2);
            var midPoint2 = q2.GetMidColinearPoints(p1, p2);

            if (((midPoint1 != null) && (Math.Round(midPoint1.DistanceTo(q1), 2) == 0) && (midPoint2 != null) && ((Math.Round(midPoint2.DistanceTo(q2), 2) == 0)))
                || ((midPoint1 != null) && (Math.Round(midPoint1.DistanceTo(q1), 2) == 0) && (Math.Round(q2.DistanceTo(p1), 2) == 0))
                || ((midPoint1 != null) && (Math.Round(midPoint1.DistanceTo(q1), 2) == 0) && (Math.Round(q2.DistanceTo(p2), 2) == 0))
                || ((midPoint2 != null) && (Math.Round(midPoint2.DistanceTo(q2), 2) == 0) && (Math.Round(q1.DistanceTo(p1), 2) == 0))
                || ((midPoint2 != null) && (Math.Round(midPoint2.DistanceTo(q2), 2) == 0) && (Math.Round(q1.DistanceTo(p2), 2) == 0)))
            {
                isChildLine = true;
            }
            return isChildLine;
        }
        public static bool CheckOneLineSameToAnother(this Line line1, Line line2)
        {
            if ((line1 is null) || (line2 is null))
            {
                return false;
            }
            bool same = false;
            var p1 = line1.GetEndPoint(0);
            var p2 = line1.GetEndPoint(1);
            var q1 = line2.GetEndPoint(0);
            var q2 = line2.GetEndPoint(1);
            if ((Math.Round(p1.DistanceTo(q1), 2) == 0) && (Math.Round(p2.DistanceTo(q2), 2) == 0) 
                || (Math.Round(p1.DistanceTo(q2), 2) == 0) && (Math.Round(p2.DistanceTo(q1), 2) == 0))
            {
                same = true;
            }
            return same;
        }
        public static Line MoveLineToLevel(this CurveElement curve, double z)
        {
            var lineStartPoint = curve.GeometryCurve.GetEndPoint(0);
            var lineEndPoint = curve.GeometryCurve.GetEndPoint(1);
            var newLineStartPoint = new XYZ(lineStartPoint.X, lineStartPoint.Y, z);
            var newLineEndPoint = new XYZ(lineEndPoint.X, lineEndPoint.Y, z);
            var newLine = Line.CreateBound(newLineStartPoint, newLineEndPoint);
            return newLine;
        }
        public static List<XYZ> GetCurveIntersectPoints(this Curve comparedCurve, List<Curve> curveList)
        {
            if ((comparedCurve == null)
                || (curveList == null)
                || (curveList.Count == 0))
            {
                return null;
            }
            List<XYZ> xPointList = new List<XYZ>();
            foreach (var curve in curveList)
            {
                IntersectionResultArray xPoint;
                SetComparisonResult xResult = comparedCurve.Intersect(curve, out xPoint);
                if (xResult == SetComparisonResult.Overlap)
                {
                    xPointList.Add(xPoint.get_Item(0).XYZPoint);
                }
            }
            return xPointList;
        }
        public static List<XYZ> GetCurveElementIntersectPoints(this CurveElement comparedCurveEle, List<CurveElement> curveEleList)
        {
            if ((comparedCurveEle == null)
                || (curveEleList == null)
                || (curveEleList.Count == 0))
            {
                return null;
            }
            Curve comparedCurve = comparedCurveEle.GeometryCurve;
            List<Curve> curveList = new List<Curve>();
            curveEleList.ForEach(c => curveList.Add(c.GeometryCurve));
            return comparedCurve.GetCurveIntersectPoints(curveList);
        }
        public static (XYZ, XYZ) GetTwoNearestPointsInBothSide(this CurveElement curve, XYZ comparedPoint, List<XYZ> xPointList)
        {
            if ((curve == null) || (comparedPoint == null) || (xPointList == null) || (xPointList.Count == 0))
            {
                return (null, null);
            }
            var curveStartPoint = curve.GeometryCurve.GetEndPoint(0);
            var curveEndPoint = curve.GeometryCurve.GetEndPoint(1);
            var newCurveStartPoint = new XYZ(curveStartPoint.X, curveStartPoint.Y, comparedPoint.Z);
            var newCurveEndPoint = new XYZ(curveEndPoint.X, curveEndPoint.Y, comparedPoint.Z);
            var comparedPointBelongToCurve = comparedPoint.CheckColinearPoints(newCurveStartPoint, newCurveEndPoint);
            if (comparedPoint.CheckCoincidentPoints(newCurveStartPoint))
            {
                comparedPointBelongToCurve = true;
            }
            if (comparedPoint.CheckCoincidentPoints(newCurveEndPoint))
            {
                comparedPointBelongToCurve = true;
            }
            if ((comparedPointBelongToCurve == false)
                || (xPointList == null) 
                || (xPointList.Count == 0))
            {
                return (null, null);
            }
            XYZ startVec = newCurveStartPoint - comparedPoint;
            XYZ endVec = newCurveEndPoint - comparedPoint;
            double minStartDist = double.MaxValue;
            double minEndDist = double.MaxValue;
            XYZ minStartPoint = null;
            XYZ minEndPoint = null;
            
            foreach (var xPoint in xPointList)
            {
                var newXPoint = new XYZ(xPoint.X, xPoint.Y, comparedPoint.Z);
                XYZ curVec = newXPoint - comparedPoint;
                if (!(curVec.CheckCoincidentPoints(XYZ.Zero))
                    && !(startVec.CheckCoincidentPoints(XYZ.Zero)) 
                    && (Math.Round(curVec.AngleTo(startVec), 3) == 0))
                {
                    var startDist = curVec.GetLength();
                    if (minStartDist > startDist)
                    {
                        minStartDist = startDist;
                        minStartPoint = xPoint;
                    }
                }
                else if (!(curVec.CheckCoincidentPoints(XYZ.Zero))
                        && !(endVec.CheckCoincidentPoints(XYZ.Zero))
                        && (Math.Round(curVec.AngleTo(endVec), 3) == 0))
                {
                    var endDist = curVec.GetLength();
                    if (minEndDist > endDist)
                    {
                        minEndDist = endDist;
                        minEndPoint = xPoint;
                    }
                }
            }

            if (minStartPoint == null)
            {
                minStartPoint = curveStartPoint;
            }
            if (minEndPoint == null)
            {
                minEndPoint = curveEndPoint;
            }
            return (minStartPoint, minEndPoint);
        }
        public static List<CurveElement> GetCurvesThroughPoint(this XYZ point, List<CurveElement> curveList)
        {
            if ((point == null) 
                || (curveList == null) 
                || (curveList.Count == 0))
            {
                return null;
            }
            List<CurveElement> xCurveList = new List<CurveElement>();
            foreach (var curve in curveList)
            {
                var curveStartPoint = curve.GeometryCurve.GetEndPoint(0);
                var curveEndPoint = curve.GeometryCurve.GetEndPoint(1);
                if ((point.CheckCoincidentPoints(curveStartPoint))
                    || (point.CheckCoincidentPoints(curveEndPoint)))
                {
                    xCurveList.Add(curve);
                }
                else if (point.CheckColinearPoints(curveStartPoint, curveEndPoint))
                {
                    var midPoint = point.GetMidColinearPoints(curveStartPoint, curveEndPoint);
                    if (Math.Round(point.DistanceTo(midPoint), 3) == 0)
                    {
                        xCurveList.Add(curve);
                    }
                }
            }
            return xCurveList;
        }
        public static XYZ GetNearestVectorFromLeftToRight (this XYZ comparedVec, List<XYZ> vecList)
        {
            if ((comparedVec == null) 
                || (vecList == null)
                || (vecList.Count == 0))
            {
                return null;
            }
            return null;
        }
        public static int IsVectorBelowAnother (this XYZ baseVec, XYZ comparedVec)
        {
            Transform tf90 = Transform.CreateRotation(XYZ.BasisZ, 90 * Math.PI / 180);
            XYZ yDownAxis = tf90.OfVector(baseVec);
            double length = comparedVec.DotProduct(yDownAxis);
            if (Math.Round(length, 3) > 0)
            {
                return 1;
            }
            else if (Math.Round(length, 3) < 0)
            {
                return -1;
            }
            else
            {
                return 0;
            }
        }
        public static bool IsVectorInList(this XYZ baseVec, List<XYZ> vecList)
        {
            if ((baseVec == null) 
                || (vecList == null) 
                || (vecList.Count == 0))
            {
                return false;
            }
            var negateVec = baseVec.Negate();
            foreach (var vec in vecList)
            {
                if ((baseVec.CheckCoincidentPoints(vec))
                    || (negateVec.CheckCoincidentPoints(vec)))
                {
                    return true;
                }
            }
            return false;
        }
        public class CurveVec
        {
            public CurveElement Curve { get; set; }
            public XYZ Vector { get; set; }
        }
        public static bool HasInverseCurveVecInList(this CurveVec baseCv, List<CurveVec> cvList)
        {
            if ((baseCv == null)
                || (cvList == null)
                || (cvList.Count == 0))
            {
                return false;
            }

            foreach (var cv in cvList)
            {
                if (baseCv.Curve.CheckSameTwoCurves(cv.Curve))
                {
                    var angle = baseCv.Vector.AngleTo(cv.Vector);
                    if (Math.Round(angle,3) == Math.Round(Math.PI, 3))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        public static XYZ GetTopLeftPointOfCurveLoop(this CurveLoop curveLoop)
        {
            if (curveLoop == null)
            {
                return null;
            }
            XYZ topLeftPoint = null;
            foreach (var curve in curveLoop)
            {
                var p1 = curve.GetEndPoint(0);
                var p2 = curve.GetEndPoint(1);
                XYZ tempTopLeftPoint = new XYZ();
                if (p1.X < p2.X)
                {
                    tempTopLeftPoint = p1;
                }
                else if (Math.Round(p1.X, 3) == Math.Round(p2.X, 3))
                {
                    if (p1.Y > p2.Y)
                    {
                        tempTopLeftPoint = p1;
                    }
                    else
                    {
                        tempTopLeftPoint = p2;
                    }
                }
                else
                {
                    tempTopLeftPoint = p2;
                }

                if (topLeftPoint == null)
                {
                    topLeftPoint = tempTopLeftPoint;
                }
                else
                {
                    if (topLeftPoint.X > tempTopLeftPoint.X)
                    {
                        topLeftPoint = tempTopLeftPoint;
                    }
                    else if (Math.Round(topLeftPoint.X, 3) == Math.Round(tempTopLeftPoint.X, 3))
                    {
                        if (topLeftPoint.Y < tempTopLeftPoint.Y)
                        {
                            topLeftPoint = tempTopLeftPoint;
                        }
                    }
                }
            }
            return topLeftPoint;
        }
       
        public static CurveLoop GetBoundFromPoint(this XYZ loc, List<CurveElement> lineColl)
        {
            CurveLoop curveLoop = null;

            var leftPoint = new XYZ(loc.X - 10 / 0.3048, loc.Y, loc.Z);
            var leftLine = Line.CreateBound(loc, leftPoint);

            #region Get First Bound
            CurveElement firstBound = null;
            XYZ firstXPoint = null;
            double firstMinDist = double.MaxValue;
            foreach (var line in lineColl)
            {
                var glLine = line.MoveLineToLevel(loc.Z);
                IntersectionResultArray xPoint;
                SetComparisonResult xResult = leftLine.Intersect(glLine, out xPoint);
                if (xResult == SetComparisonResult.Overlap)
                {
                    var dist = loc.DistanceTo(xPoint.get_Item(0).XYZPoint);
                    if (firstMinDist > dist)
                    {
                        firstMinDist = dist;
                        firstXPoint = xPoint.get_Item(0).XYZPoint;
                        firstBound = line;
                    }
                }
            }
            #endregion

            #region Define Continue Point
            List<XYZ> firstBoundXPointList = firstBound.GetCurveElementIntersectPoints(lineColl);
            XYZ minStartPoint = firstBound.GetTwoNearestPointsInBothSide(firstXPoint, firstBoundXPointList).Item1;
            XYZ minEndPoint = firstBound.GetTwoNearestPointsInBothSide(firstXPoint, firstBoundXPointList).Item2;
            XYZ continuePoint = null;
            XYZ closePoint = null;
            if (minStartPoint.Y > minEndPoint.Y)
            {
                continuePoint = minStartPoint;
                closePoint = minEndPoint;
            }
            else
            {
                continuePoint = minEndPoint;
                closePoint = minStartPoint;
            }
            XYZ baseVec = loc - continuePoint;
            #endregion Define Continue Point

            List<XYZ> boundPointList = new List<XYZ>();
            boundPointList.Add(continuePoint);
            bool endLoop = false;
            CurveElement curCurve = firstBound;
            List<CurveVec> contCvList = new List<CurveVec>();
            while (endLoop == false)
            {
                //var newLineColl = lineColl.RemoveCurveInCurveList(curCurve);
                List<CurveElement> xCurveList = continuePoint.GetCurvesThroughPoint(lineColl);
                CurveElement continueCurve = null;
                XYZ continueVec = null;
                if (xCurveList.Count > 0)
                {
                    double minAngle = double.MaxValue;
                    foreach (var xCurve in xCurveList)
                    {
                        var xCurveStartPoint = xCurve.GeometryCurve.GetEndPoint(0);
                        var xCurveEndPoint = xCurve.GeometryCurve.GetEndPoint(1);
                        XYZ xStartVec = xCurveStartPoint - continuePoint;
                        XYZ xEndVec = xCurveEndPoint - continuePoint;
                        int isStartVecBelowBaseVec = baseVec.IsVectorBelowAnother(xStartVec);
                        int isEndVecBelowBaseVec = baseVec.IsVectorBelowAnother(xEndVec);
                        XYZ selectedVec = null;
                        bool isAllBelow = false;
                        if ((isStartVecBelowBaseVec == 1)
                            && (!xStartVec.CheckCoincidentPoints(XYZ.Zero)))
                        {
                            selectedVec = xStartVec;
                        }
                        else if ((isEndVecBelowBaseVec == 1)
                                && (!xEndVec.CheckCoincidentPoints(XYZ.Zero)))
                        {
                            selectedVec = xEndVec;
                        }
                        if (selectedVec == null)
                        {
                            isAllBelow = true;
                            if (!(xStartVec.CheckCoincidentPoints(XYZ.Zero)))
                            {
                                selectedVec = xStartVec;
                            }
                            else if (!(xEndVec.CheckCoincidentPoints(XYZ.Zero)))
                            {
                                selectedVec = xEndVec;
                            }
                        }

                        CurveVec tempCv = new CurveVec();
                        tempCv.Curve = xCurve;
                        tempCv.Vector = selectedVec;

                        if ((selectedVec != null)
                            && (!tempCv.HasInverseCurveVecInList(contCvList)))
                        {
                            double angle = baseVec.AngleTo(selectedVec);
                            if (isAllBelow == true)
                            {
                                angle = (360 * Math.PI / 180) - angle;
                            }
                            if (minAngle > angle)
                            {
                                minAngle = angle;
                                continueCurve = xCurve;
                                curCurve = continueCurve;
                                continueVec = selectedVec;
                                CurveVec cv = new CurveVec();
                                cv.Curve = continueCurve;
                                cv.Vector = continueVec;
                                contCvList.Add(cv);
                            }
                        }
                    }
                }
                #region Define Continue Point
                List<XYZ> curCurveXPointList = curCurve.GetCurveElementIntersectPoints(lineColl);
                XYZ curCurveMinStartPoint = curCurve.GetTwoNearestPointsInBothSide(continuePoint, curCurveXPointList).Item1;
                XYZ curCurveMinEndPoint = curCurve.GetTwoNearestPointsInBothSide(continuePoint, curCurveXPointList).Item2;
                XYZ contStartVec = curCurveMinStartPoint - continuePoint;
                XYZ contEndVec = curCurveMinEndPoint - continuePoint;
                double startAngle = 0.0;
                double endAngle = 0.0;
                if (contStartVec.CheckCoincidentPoints(XYZ.Zero))
                {
                    startAngle = -1;
                }
                else
                {
                    startAngle = Math.Round(continueVec.AngleTo(contStartVec), 3);
                }
                if (contEndVec.CheckCoincidentPoints(XYZ.Zero))
                {
                    endAngle = -1;
                }
                else
                {
                    endAngle = Math.Round(continueVec.AngleTo(contEndVec), 3);
                }

                if (startAngle == 0)
                {
                    continuePoint = curCurveMinStartPoint;
                }
                else if (endAngle == 0)
                {
                    continuePoint = curCurveMinEndPoint;
                }
                baseVec = loc - continuePoint;
                boundPointList.Add(continuePoint);
                #endregion Define Continue Point

                if (continuePoint.CheckCoincidentPoints(closePoint))
                {
                    endLoop = true;
                }
                bool isColinear = continuePoint.CheckColinearPoints(minStartPoint, minEndPoint);
                if (isColinear)
                {
                    XYZ midColinearPoint = continuePoint.GetMidColinearPoints(minStartPoint, minEndPoint);
                    bool isContinuePointInMid = (continuePoint.CheckCoincidentPoints(midColinearPoint)) ? true : false;
                    if (isContinuePointInMid == true)
                    {
                        endLoop = true;
                    }
                }

                List<Curve> boundCurveList = new List<Curve>();
                if (boundPointList.Count > 2)
                {
                    for (int i = 0; i < boundPointList.Count - 1; i++)
                    {
                        var curPoint = boundPointList.ElementAt(i);
                        var nextPoint = boundPointList.ElementAt(i + 1);
                        Curve curve = Line.CreateBound(curPoint, nextPoint);
                        boundCurveList.Add(curve);
                    }
                    Curve lastCurve = Line.CreateBound(boundPointList.ElementAt(boundPointList.Count - 1), boundPointList.ElementAt(0));
                    boundCurveList.Add(lastCurve);
                    curveLoop = CurveLoop.Create(boundCurveList);
                }
            }

            return curveLoop;
        }
        public static bool LineInsideFace(this Line line, Face face, bool countEdge = true)
        {
            if ((line == null) || (face == null))
            {
                return false;
            }
            Plane plane = face.GetSurface() as Plane;
            var projectLine = line.ProjectLineToPlane(plane);
            UV uv1, uv2 = new UV();
            plane.Project(line.GetEndPoint(0), out uv1, out double dist1);
            plane.Project(line.GetEndPoint(1), out uv2, out double dist2);
            bool uv1InsideFace = face.IsInside(uv1);
            bool uv2InsideFace = face.IsInside(uv2);
            
            bool onEdge = false;
            var curveLoops = face.GetEdgesAsCurveLoops();
            foreach (var curveLoop in curveLoops)
            {
                foreach (Line faceLine in curveLoop)
                {
                    onEdge = faceLine.CheckOneLineBelongToAnother(line);
                    if (onEdge) break;
                }
                if (onEdge) break;
            }
            if ((uv1InsideFace == true) 
                && (uv2InsideFace == true))
            {
                if (!onEdge)
                {
                    return true;
                }
                else
                {
                    if (countEdge)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return false;
        }
        public static bool LineInsideFaces(this Line line, List<Face> faces, bool countEdge = true)
        {
            if ((line == null) || (faces == null) || (faces.Count == 0))
            {
                return false;
            }
            bool onEdge = false;
            foreach (var face in faces)
            {
                var curveLoops = face.GetEdgesAsCurveLoops();
                foreach (var curveLoop in curveLoops)
                {
                    foreach (Line faceLine in curveLoop)
                    {
                        onEdge = faceLine.CheckOneLineBelongToAnother(line);
                        if (onEdge) break;
                    }
                    if (onEdge) break;
                }
                if ((line.LineInsideFace(face, countEdge)) && (!onEdge))
                {
                    return true;
                }
                else if ((line.LineInsideFace(face, countEdge)) && (onEdge))
                {
                    if (countEdge) return true;
                    else return false;
                }
            }
            return false;
        }
        public static bool LineInsideCurveLoop(this Line line, CurveLoop curveLoop)
        {
            if ((line == null) || (curveLoop == null))
            {
                return false;
            }

            List<CurveLoop> curveLoopList = new List<CurveLoop>();
            curveLoopList.Add(curveLoop);
            Solid curveLoopExtrude = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoopList, XYZ.BasisZ, 10);
            Face face = curveLoopExtrude.GetBottomPlanarFace();
            return line.LineInsideFace(face);
        }
        public static (int, List<XYZ>) IntersectBwLineAndCl(this Line line, CurveLoop curveLoop)
        {
            if ((line == null) || (curveLoop == null))
            {
                return (0, null);
            }
            XYZ p1 = line.GetEndPoint(0);
            XYZ p2 = line.GetEndPoint(1);
            XYZ newP1 = new XYZ(p1.X, p1.Y, 0);
            XYZ newP2 = new XYZ(p2.X, p2.Y, 0);
            line = Line.CreateBound(newP1, newP2);
            int noOfIntersect = 0;
            List<XYZ> intersectPointList = new List<XYZ>();
            foreach (var curve in curveLoop)
            {
                XYZ cp1 = curve.GetEndPoint(0);
                XYZ cp2 = curve.GetEndPoint(1);
                XYZ cnewP1 = new XYZ(cp1.X, cp1.Y, 0);
                XYZ cnewP2 = new XYZ(cp2.X, cp2.Y, 0);
                var newCurve = Line.CreateBound(cnewP1, cnewP2);

                IntersectionResultArray xPoint;
                SetComparisonResult xResult = newCurve.Intersect(line, out xPoint);
                if (xResult == SetComparisonResult.Overlap)
                {
                    noOfIntersect++;
                    intersectPointList.Add(xPoint.get_Item(0).XYZPoint);
                }
            }
            return (noOfIntersect, intersectPointList);
        }
        public static (int, List<XYZ>, CurveLoop) IntersectBwLineAndCls(this Line line, List<CurveLoop> curveLoopList)
        {
            if ((line == null) 
                || (curveLoopList == null)
                || (curveLoopList.Count == 0))
            {
                return (0, null, null);
            }
            XYZ p1 = line.GetEndPoint(0);
            XYZ p2 = line.GetEndPoint(1);
            XYZ newP1 = new XYZ(p1.X, p1.Y, 0);
            XYZ newP2 = new XYZ(p2.X, p2.Y, 0);
            line = Line.CreateBound(newP1, newP2);
            int noOfIntersect = 0;
            List<XYZ> intersectPointList = new List<XYZ>();
            CurveLoop intersectCurveLoop = null;
            foreach (var curveLoop in curveLoopList)
            {
                var xResult = line.IntersectBwLineAndCl(curveLoop);
                var tempNoOfIntersect = xResult.Item1;
                var tempIntersectPointList = xResult.Item2;
                if ((tempNoOfIntersect > 0) 
                    && (tempIntersectPointList != null)
                    && (tempIntersectPointList.Count > 0))
                {
                    noOfIntersect += tempNoOfIntersect;
                    tempIntersectPointList.ForEach(p => intersectPointList.Add(p));
                    intersectCurveLoop = curveLoop;
                }
            }
            return (noOfIntersect, intersectPointList, intersectCurveLoop);
        }
        public static List<Line> LineAfterIntersectCurveLoop(this Line line, CurveLoop curveLoop, bool getExLines = true)
        {
            if ((line == null) || (curveLoop == null))
            {
                return null;
            }
            XYZ p1 = line.GetEndPoint(0);
            XYZ p2 = line.GetEndPoint(1);
            XYZ newP1 = new XYZ(p1.X, p1.Y, 0);
            XYZ newP2 = new XYZ(p2.X, p2.Y, 0);
            line = Line.CreateBound(newP1, newP2);

            List<Line> exLineList = new List<Line>();
            List<Line> inLineList = new List<Line>();
            var lineIntersectCurveLoop = line.IntersectBwLineAndCl(curveLoop);
            int noOfIntersect = lineIntersectCurveLoop.Item1;
            List<XYZ> intersectList = lineIntersectCurveLoop.Item2;
            double length = line.Length;

            switch (noOfIntersect)
            {
                case 1:
                    var intersectPoint = intersectList[0];

                    Line line1 = null;
                    Line line2 = null;
                    if (!newP1.CheckCoincidentPoints(intersectPoint))
                    {
                        line1 = Line.CreateBound(newP1, intersectPoint);
                    }
                    if (!newP2.CheckCoincidentPoints(intersectPoint))
                    {
                        line2 = Line.CreateBound(newP2, intersectPoint);
                    }

                    if ((line1 != null) && (line1.LineInsideCurveLoop(curveLoop)))
                    {
                        exLineList.Add(line2);
                        inLineList.Add(line1);
                    }
                    else if ((line2 != null) && (line2.LineInsideCurveLoop(curveLoop)))
                    {
                        exLineList.Add(line1);
                        inLineList.Add(line2);
                    }

                    break;
                case 2:
                    XYZ q1 = intersectList[0];
                    XYZ q2 = intersectList[1];

                    var midLine = Line.CreateBound(q1, q2);
                    inLineList.Add(midLine);

                    Line line11 = null;
                    Line line12 = null;
                    Line line21 = null;
                    Line line22 = null;

                    double length11 = 0.0;
                    double length12 = 0.0;
                    double length21 = 0.0;
                    double length22 = 0.0;

                    if (newP1.CheckCoincidentPoints(q1) == false)
                    {
                        line11 = Line.CreateBound(newP1, q1);
                        length11 = line11.Length;
                    }
                    if (newP2.CheckCoincidentPoints(q2) == false)
                    {
                        line12 = Line.CreateBound(newP2, q2);
                        length12 = line12.Length;
                    }

                    double length1 = length11 + length12;

                    if (newP1.CheckCoincidentPoints(q2) == false)
                    {
                        line21 = Line.CreateBound(newP1, q2);
                        length21 = line21.Length;
                    }
                    if (newP2.CheckCoincidentPoints(q1) == false)
                    {
                        line22 = Line.CreateBound(newP2, q1);
                        length22 = line22.Length;
                    }
                    
                    double length2 = length21 + length22;

                    if ((length1 < length) && (length2 > length))
                    {
                        if (line11 != null)
                        {
                            exLineList.Add(line11);
                        }
                        if (line12 != null)
                        {
                            exLineList.Add(line12);
                        }
                    }
                    else if ((length1 > length) && (length2 < length))
                    {
                        if (line21 != null)
                        {
                            exLineList.Add(line21);
                        }
                        if (line22 != null)
                        {
                            exLineList.Add(line22);
                        }
                    }
                    break;
                default:
                    break;
            }
            if (exLineList.Count == 0)
            {
                exLineList.Add(line);
            }
            if (inLineList.Count == 0)
            {
                inLineList.Add(line);
            }
            if (getExLines)
            {
                return exLineList;
            }
            else
            {
                return inLineList;
            }
        }
        public static List<Line> LineAfterIntersectCurveLoops(this Line line, List<CurveLoop> curveLoopList, bool getExLines = true)
        {
            if ((line == null) 
                || (curveLoopList == null)
                || (curveLoopList.Count == 0))
            {
                return null;
            }
            List<Line> listOfLineAfterIntersect = new List<Line>();
            var lineIntersectCurveLoops = line.IntersectBwLineAndCls(curveLoopList);
            int noOfIntersect = lineIntersectCurveLoops.Item1;
            List<XYZ> intersectList = lineIntersectCurveLoops.Item2;
            CurveLoop intersectCurveLoop = lineIntersectCurveLoops.Item3;
            if (intersectCurveLoop != null)
            {
                listOfLineAfterIntersect = line.LineAfterIntersectCurveLoop(intersectCurveLoop, getExLines);
            }
            else
            {
                listOfLineAfterIntersect.Add(line);
            }
            return listOfLineAfterIntersect;
        }
        public static Line GetCenterLineOfCurveLoop(this CurveLoop cl, Application app, Document doc, double fkThickness)
        {
            if (cl is null)
            {
                return null;
            }
            List<Curve> curveList = new List<Curve>();
            foreach (var curve in cl)
            {
                if (Math.Round(curve.Length - fkThickness, 3) <= 5 / 304.8)
                {
                    curveList.Add(curve);
                }
            }
            if (curveList.Count == 2)
            {
                var curve1 = curveList[0];
                var curve2 = curveList[1];
                var p1 = curve1.GetEndPoint(0);
                var p2 = curve1.GetEndPoint(1);
                var p3 = curve2.GetEndPoint(0);
                var p4 = curve2.GetEndPoint(1);
                var midPoint1 = (p1 + p2) / 2;
                var midPoint2 = (p3 + p4) / 2;
                var centerLine = Line.CreateBound(midPoint1, midPoint2);
                return centerLine;
            }
            else
            {
                return null;
            }
        }
        public static CurveLoop RefineCurveLoop(this CurveLoop curveLoop)
        {
            CurveLoop refinedCurveLoop = new CurveLoop();
            int i = 0;
            while (i < curveLoop.Count() - 1)
            {
                var curCurve = curveLoop.ElementAt(i) as Line;
                var curDir = curCurve.Direction;
                var nextCurve = curveLoop.ElementAt(i + 1) as Line;
                var nextDir = nextCurve.Direction;
                if (curDir.AngleTo(nextDir) != 0)
                {
                    refinedCurveLoop.Append(curCurve);
                    i++;
                }
                else
                {
                    var newCurve = Line.CreateBound(curCurve.GetEndPoint(0), nextCurve.GetEndPoint(1));
                    refinedCurveLoop.Append(newCurve);
                    i += 2;
                }
            }
            refinedCurveLoop.Append(curveLoop.ElementAt(curveLoop.Count() - 1));
            return refinedCurveLoop;
        }
        public static bool AreTwoLinesSameDir(this Line line1, Line line2)
        {
            if ((line1 == null) || (line2 == null))
            {
                return false;
            }
            var line1Dir = line1.Direction;
            var line2Dir = line2.Direction;
            if ((Math.Round(line1Dir.AngleTo(line2Dir), 3) == 0) || (Math.Round(line1Dir.AngleTo(line2Dir) - Math.PI, 3) == 0))
            {
                return true;
            }
            else return false;
        }
        public static (bool, XYZ) DoTwoLinesShareSamePoint(this Line line1, Line line2)
        {
            if ((line1 == null) || (line2 == null))
            {
                return (false, null);
            }
            var p1 = line1.GetEndPoint(0);
            var p2 = line1.GetEndPoint(1);
            var p3 = line2.GetEndPoint(0);
            var p4 = line2.GetEndPoint(1);

            if (line1.CheckTwoCurvesIfTheyAreSame(line2) == false)
            {
                if ((p1.CheckCoincidentPoints(p3))
                || (p1.CheckCoincidentPoints(p4)))
                {
                    return (true, p1);
                }
                else if ((p2.CheckCoincidentPoints(p3))
                    || (p2.CheckCoincidentPoints(p4)))
                {
                    return (true, p2);
                }
                else return (false, null);
            }
            return (false, null);
        }
        public static bool AreTwoLinesContinuous(this Line line1, Line line2)
        {
            if ((line1 == null) || (line2 == null))
            {
                return false;
            }
            if ((line1.AreTwoLinesSameDir(line2))
                && (line1.DoTwoLinesShareSamePoint(line2)).Item1)
            {
                return true;
            }
            else return false;
        }
        public static Line CombineContinuousLines(this Line line1, Line line2)
        {
            if ((line1 == null) || (line2 == null))
            {
                return null;
            }
            if (line1.AreTwoLinesContinuous(line2))
            {
                var sharedPoint = line1.DoTwoLinesShareSamePoint(line2).Item2;
                var p1 = (line1.GetEndPoint(0).CheckCoincidentPoints(sharedPoint)) ? line1.GetEndPoint(1) : line1.GetEndPoint(0);
                var p2 = (line2.GetEndPoint(0).CheckCoincidentPoints(sharedPoint)) ? line2.GetEndPoint(1) : line2.GetEndPoint(0);
                return Line.CreateBound(p1, p2);
            }
            else return null;
        }
        public static bool IsXVec(this XYZ vec)
        {
            if (vec is null)
            {
                return false;
            }
            double angle = vec.AngleTo(XYZ.BasisX);
            if (angle < Math.PI / 4 || angle > Math.PI * 3 / 4)
            {
                return true;
            }
            else return false;
        }
        public static XYZ GetCurveMid(this Curve curve)
        {
            if (curve is null)
            {
                return null;
            }
            XYZ p1 = curve.GetEndPoint(0);
            XYZ p2 = curve.GetEndPoint(1);
            XYZ mid = new XYZ((p1.X + p2.X) / 2, (p1.Y + p2.Y) / 2, (p1.Z + p2.Z) / 2);
            return mid;
        }
    }
}
