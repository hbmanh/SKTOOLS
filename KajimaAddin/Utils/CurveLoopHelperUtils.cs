using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.IFC;

namespace SKToolsAddins.Utils
{
    public static class CurveLoopHelperUtils
    {
        public static CurveLoop AppendFrom(this CurveLoop curveLoop, ModelCurveArray mca)
        {
            List<ModelCurve> modelCurveList = new List<ModelCurve>();
            foreach (ModelCurve m in mca)
            {
                modelCurveList.Add(m);
            }
            TryToAppend(curveLoop, modelCurveList[0]);
            modelCurveList.RemoveAt(0);
            while (modelCurveList.Count > 0)
            {
                foreach (var item in modelCurveList)
                {
                    if (TryToAppend(curveLoop, item))
                    {
                        modelCurveList.Remove(item);
                        break;
                    }
                }
            }
            return null;
        }
        private static bool TryToAppend(CurveLoop curveLoop, ModelCurve modelLine)
        {
            try
            {
                curveLoop.Append(modelLine.GeometryCurve);
                return true;
            }
            catch (Exception)
            {
                Debug.WriteLine("Append Curve Failed!");
            }
            try
            {
                curveLoop.Append(modelLine.GeometryCurve.CreateReversed());
                return true;
            }
            catch (Exception)
            {
            }
            return false;
        }
        public static CurveLoop AppendFromCurveArray(this CurveLoop curveLoop, CurveArray ca)
        {
            List<Curve> curveList = new List<Curve>();
            foreach (Curve m in ca)
            {
                curveList.Add(m);
            }
            TryToAppendCurve(curveLoop, curveList[0]);
            curveList.RemoveAt(0);
            int i = 0;
            while (curveList.Count > 0)
            {
                if (i > 100)
                {
                    break;
                }
                foreach (var item in curveList)
                {
                    if (TryToAppendCurve(curveLoop, item))
                    {
                        curveList.Remove(item);
                        i = 0;
                        break;
                    }
                }
                i++;
            }
            return null;
        }
        private static bool TryToAppendCurve(CurveLoop curveLoop, Curve curve)
        {
            try
            {
                curveLoop.Append(curve);
                return true;
            }
            catch (Exception)
            {
            }
            try
            {
                curveLoop.Append(curve.CreateReversed());
                return true;
            }
            catch (Exception)
            {
            }
            return false;
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
        public static CurveLoop GetMaxCurveLoop(this IList<CurveLoop> curveLoops)
        {
            CurveLoop maxCurveLoop = null;
            if ((curveLoops is null) || (curveLoops.Count.Equals(0)))
            {
                return null;
            }
            else
            {
                maxCurveLoop = curveLoops[0];
                List<CurveLoop> maxCurveLoopList = new List<CurveLoop>();
                maxCurveLoopList.Add(maxCurveLoop);
                var maxCurveLoopArea = ExporterIFCUtils.ComputeAreaOfCurveLoops(maxCurveLoopList);
                if (curveLoops.Count > 1)
                {
                    for (int i = 1; i < curveLoops.Count; i++)
                    {
                        var curCurveLoop = curveLoops[i];
                        List<CurveLoop> curCurveLoopList = new List<CurveLoop>();
                        curCurveLoopList.Add(curCurveLoop);
                        var curCurveLoopArea = ExporterIFCUtils.ComputeAreaOfCurveLoops(curCurveLoopList);
                        if (maxCurveLoopArea < curCurveLoopArea)
                        {
                            maxCurveLoop = curCurveLoop;
                            maxCurveLoopArea = curCurveLoopArea;
                        }
                    }
                }
                return maxCurveLoop;
            }
        }
    }
}
