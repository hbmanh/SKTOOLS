using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace SKRevitAddins.Utils
{
    public static class ListUtils
    {
        private static Random rng = new Random();
        public static void Shuffle<T>(this IList<T> list)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rng.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }
        public static CurveLoop ShiftCurveLoopRandom(this CurveLoop curveloop)
        {
            CurveLoop newCurveLoop = new CurveLoop();
            Random rnd = new Random();
            int shift = rnd.Next(1, curveloop.Count());
            for (int i = shift; i < curveloop.Count(); i++)
            {
                Curve curve = curveloop.ElementAt(i);
                newCurveLoop.Append(curve);
            }
            for (int i = 0; i < shift; i++)
            {
                Curve curve = curveloop.ElementAt(i);
                newCurveLoop.Append(curve);
            }
            return newCurveLoop;
        }
    }
}
