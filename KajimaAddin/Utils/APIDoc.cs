using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SKToolsAddins
{
    public static class APIDoc
    {
        //Selection
        public class Selection
        {
            private List<ElementId> _ids;

            //public Selection()
            //{
            //    _ids = new List<ElementId>();
            //}

            public void Add(ElementId id)
            {
                _ids.Add(id);
            }

            public void Remove(ElementId id)
            {
                _ids.Remove(id);
            }

            public void Clear()
            {
                _ids.Clear();
            }

            public ICollection<ElementId> GetElementIds()
            {
                return _ids;
            }
        }
        public static double CPGridLength { get { return MmToFeet(12000); } } // Declare variable CPGridLength

        public static double CPExtLength { get { return MmToFeet(500); } } // Declare variable CPExtLength

        

        // Create grids function
        public static void CreateGrids(Document doc, XYZ basePoint, List<double> spaces, string startName, bool isHorizontal)
        {

            using (Transaction transaction = new Transaction(doc, "Create Grids"))
            {
                transaction.Start();

                // Get grids in project
                List<ElementId> grids = new FilteredElementCollector(doc)
                .OfClass(typeof(Grid))
                .Cast<Grid>()
                .Select(x => x.Id)
                .ToList();

                // Check grid already exists in project and delete
                if (isHorizontal)
                {
                    foreach (var item in grids)
                    {
                        doc.Delete(item);
                    }
                }

                // Check the direction to create
                XYZ offsetDir = isHorizontal ? XYZ.BasisX : XYZ.BasisY;

                // Check the direction and set value of start point
                XYZ startPoint = isHorizontal ? basePoint.Add(XYZ.BasisY.Multiply(-CPExtLength)) : basePoint.Add(XYZ.BasisX.Multiply(-CPExtLength));

                // Check the direction and set value of end point
                XYZ endPoint = isHorizontal ? startPoint.Add(XYZ.BasisY.Multiply(CPGridLength + CPExtLength)) : startPoint.Add(XYZ.BasisX.Multiply(CPGridLength + CPExtLength));

                // Create line for grid
                Autodesk.Revit.DB.Line geoLine = Autodesk.Revit.DB.Line.CreateBound(startPoint, endPoint);

                // Create grid
                Autodesk.Revit.DB.Grid grid = Autodesk.Revit.DB.Grid.Create(doc, geoLine);

                // Set name of grid
                grid.Name = startName;

                foreach (double space in spaces)
                {
                    // Set value of start point for next grid
                    startPoint = startPoint.Add(offsetDir.Multiply(space));

                    // Set value of end point for next grid
                    endPoint = endPoint.Add(offsetDir.Multiply(space));

                    // Create line next grid
                    geoLine = Autodesk.Revit.DB.Line.CreateBound(startPoint, endPoint);

                    // Create next grid
                    Autodesk.Revit.DB.Grid.Create(doc, geoLine);
                }
                transaction.Commit();
            }
        }

        // Get spaces function
        public static List<double> ParseSpaces(string value)
        {
            // Check null or emplty spaces input
            if (string.IsNullOrEmpty(value))
            {
                throw new Exception("Invalid spaces input!");
            }

            // Split sign [ , ] and get spaces
            string[] sps = value.Split(',');
            List<double> spaces = new List<double>();
            foreach (string s in sps)
            {
                double val = Convert.ToDouble(s);

                // Check spaces after convert to double
                if (val <= 0)
                {
                    throw new Exception("Non positive space!");
                }

                // Add spaces to list
                spaces.Add(MmToFeet(val));
            }
            return spaces;
        }

        private const double ConvertFeetToMillimeters = 12 * 25.4; // Convert Feet To Millimeters

        // Convert Millimeters To Feet 
        public static double MmToFeet(double mm)
        {
            return mm / ConvertFeetToMillimeters;
        }

        // Create class GridFilter for filter grid
        public class GridFilter : ISelectionFilter
        {
            Document doc = null;

            // Initialization function
            public GridFilter(Document doc)
            {
                this.doc = doc;
            }

            // Method get allow element in filter
            public bool AllowElement(Element elem)
            {
                if (elem.Category.Id.IntegerValue.Equals(-2000220))
                { return true; }
                return false;
            }

            // Method get allow reference in filter
            public bool AllowReference(Reference reference, XYZ position)
            {

                if (doc.GetElement(reference) is Autodesk.Revit.DB.Grid)
                {
                    return true;
                }
                return false;
            }
        }

        // Create class  for project
        internal class SKTools
        {
            internal Grid Grid { get; set; } // Declare variable Grid
            internal string GridName { get; set; } // Declare variable GridName
            internal Reference GridReference { get { return new Reference(Grid); } } // Declare variable GridReference
            internal Line GridLine { get { return Grid.Curve as Line; } } // Declare variable GridLine

            // Initialization function
            internal SKTools(Grid grid)
            {
                this.Grid = grid;
                this.GridName = grid.Name;
            }

            // Get project base point function
            internal double DistanceToBasePoint(Document doc)
            {
                XYZ point = new XYZ(0,0,0);
                
                return GridLine.Distance(point);
            }

            // Greate dim of out grid function
            public static void DimGridOutest(IList<SKTools> grids, Document doc, Line dimLine, DimensionType dimType, View view)
            {
                //Get grid in project
                SKTools farest = grids.Aggregate((grid1, grid2) => grid1.DistanceToBasePoint(doc) > grid2.DistanceToBasePoint(doc) ? grid1 : grid2);
                SKTools nearest = grids.Aggregate((grid1, grid2) => grid1.DistanceToBasePoint(doc) < grid2.DistanceToBasePoint(doc) ? grid1 : grid2);

                //Get referncen array of grids
                ReferenceArray dimReference = new ReferenceArray();
                dimReference.Append(farest.GridReference);
                dimReference.Append(nearest.GridReference);

                //Create out grid
                Dimension dimY_Level2 = doc.Create.NewDimension(view, dimLine, dimReference, dimType);
            }
        }


    }
}
