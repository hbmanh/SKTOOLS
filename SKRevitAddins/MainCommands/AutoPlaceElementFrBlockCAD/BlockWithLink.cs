namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
{
    public class BlockWithLink
    {
        public Autodesk.Revit.DB.GeometryInstance Block { get; set; }
        public Autodesk.Revit.DB.ImportInstance CadLink { get; set; }
    }
}
