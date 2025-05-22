using Autodesk.Revit.DB;

namespace SKRevitAddins.Utils
{
    public struct PositionKey
    {
        private readonly int x, y, z;

        public PositionKey(XYZ point, double tolerance = 0.001)
        {
            x = (int)(point.X / tolerance);
            y = (int)(point.Y / tolerance);
            z = (int)(point.Z / tolerance);
        }

        public override int GetHashCode() => x ^ (y << 2) ^ (z << 4);

        public override bool Equals(object obj)
        {
            return obj is PositionKey other &&
                   x == other.x && y == other.y && z == other.z;
        }
    }
}
