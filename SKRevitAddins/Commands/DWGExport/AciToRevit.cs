using Autodesk.Revit.DB;

namespace SKRevitAddins.Commands.DWGExport
{
    internal static class ColorUtil
    {
        public static Color AciToRevit(int idx)
        {
            return idx switch
            {
                1 => new Color(255, 0, 0),
                2 => new Color(0, 255, 0),
                3 => new Color(0, 0, 255),
                4 => new Color(0, 255, 255),
                5 => new Color(255, 255, 0),
                6 => new Color(255, 0, 255),
                8 => new Color(128, 128, 128),
                9 => new Color(255, 255, 255),
                _ => new Color(0, 0, 0)
            };
        }
    }
}
