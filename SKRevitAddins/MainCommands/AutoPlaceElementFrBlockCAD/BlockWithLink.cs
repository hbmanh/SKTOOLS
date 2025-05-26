using Autodesk.Revit.DB;

namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
{
    public class BlockWithLink
    {
        public GeometryInstance Block { get; set; }
        public ImportInstance CadLink { get; set; }
    }
}
