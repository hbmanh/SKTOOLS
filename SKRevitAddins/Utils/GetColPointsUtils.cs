//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Autodesk.Revit.DB;
//using System.Linq;
//using Autodesk.Revit.UI;
//using Autodesk.Revit.ApplicationServices;
//using TakeuchiAddins.Models;
//using static TakeuchiAddins.Models.PlaxisModels;
//using System.Text.RegularExpressions;

//namespace TakeuchiAddins.Utils
//{
//    public static class GetColPointsUtils
//    {
//        public static double MatchFoundationLevel(Document doc, Element found)
//        {
//            double lowestLevel = 0.0;
//            List<FamilyInstance> jibanColl = new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType()
//                .OfCategory(BuiltInCategory.OST_CurtainWallPanels).Cast<FamilyInstance>()
//                .Where(e => e.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble() > 0).ToList();
//            var foundSolid = found.GetAllSolidsAdvance(true);
//            var foundUnionSolid = foundSolid[0];
//            for (int i = 1; i < foundSolid.Count; i++)
//            {
//                foundUnionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(foundUnionSolid, foundSolid[i], BooleanOperationsType.Union);
//            }
//            List<FamilyInstance> intersectedJibans = new List<FamilyInstance>();
//            foreach (var jiban in jibanColl)
//            {
//                var jibanSolid = jiban.GetAllSolidsAdvance(true).First();
//                var intersectedSolid = BooleanOperationsUtils.ExecuteBooleanOperation(foundUnionSolid, jibanSolid, BooleanOperationsType.Intersect);
//                if (!(intersectedSolid is null) && (Math.Round(intersectedSolid.Volume, 3) > 0))
//                {
//                    intersectedJibans.Add(jiban);
//                }
//            }
//            if (!(intersectedJibans is null) && (intersectedJibans.Count > 0))
//            {
//                lowestLevel = Math.Round(intersectedJibans.Min(e => e.Host.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM).AsDouble()) * 0.3048, 3);
//            }
//            return lowestLevel;
//        }
//        public static List<FoundGrid> GetFoundGrid(Document doc, List<FamilyInstance> foundColl)
//        {
//            BasePoint projectPoint = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).Cast<BasePoint>().First();
//            ElementId gridNoBubbleId = new ElementId(27250);
//            var gridColl = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType()
//                .Where(g => !(g.GetTypeId().Equals(gridNoBubbleId))).Cast<Grid>().ToList();
//            gridColl.Sort((a, b) => a.Name.CompareTo(b.Name));
//            List<FoundGrid> foundGrids = new List<FoundGrid>();
//            List<GridIntersection> gridIntersectionList = new List<GridIntersection>();
//            for (int i = 0; i < gridColl.Count; i++)
//            {
//                var gridA = gridColl[i];
//                var gridACurve = gridA.Curve;
//                for (int j = i + 1; j < gridColl.Count; j++)
//                {
//                    var gridB = gridColl[j];
//                    var gridBCurve = gridB.Curve;
//                    IntersectionResultArray gridIntersectResults = new IntersectionResultArray();
//                    gridACurve.Intersect(gridBCurve, out gridIntersectResults);
//                    if (gridIntersectResults != null)
//                    {
//                        GridIntersection gridIntersection = new GridIntersection();
//                        gridIntersection.GridA = gridA;
//                        gridIntersection.GridB = gridB;
//                        gridIntersection.IntersectionPoint = gridIntersectResults.get_Item(0).XYZPoint;
//                        gridIntersectionList.Add(gridIntersection);
//                        //Debug.WriteLine(gridIntersection.GridA.Name + " - " + gridIntersection.GridB.Name);
//                    }
//                }
//            }
//            if (!(foundColl == null) && (foundColl.Count > 0))
//            {
//                foreach (var found in foundColl)
//                {
//                    var foundLoc = (found.Location) as LocationPoint;
//                    var foundLocPoint = foundLoc.Point;
//                    double minDistToGrid = double.MaxValue;
//                    FoundGrid curFoundGrid = new FoundGrid();
//                    curFoundGrid.FoundInst = found;
//                    foreach (var gi in gridIntersectionList)
//                    {
//                        var dist = foundLocPoint.DistanceTo(gi.IntersectionPoint);
//                        if (minDistToGrid > dist)
//                        {
//                            minDistToGrid = dist;
//                            curFoundGrid.GridA = gi.GridA;
//                            curFoundGrid.GridB = gi.GridB;
//                            curFoundGrid.GridDist = minDistToGrid;
//                        }
//                    }
//                    foundGrids.Add(curFoundGrid);
//                }
//            }
//            return foundGrids;
//        }
//        public static (List<string>, List<UserPointLoad>) GetColPoints(Document doc, List<FamilyInstance> foundColl)
//        {
//            BasePoint projectPoint = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).Cast<BasePoint>().First();
//            List<string> colCommands = new List<string>();
//            List<UserPointLoad> UserPointLoads = new List<UserPointLoad>();
//            var foundGridList = GetFoundGrid(doc, foundColl);
//            ElementId colStyleId = new ElementId(868429);
//            var tempForceCommand = "tempForceCommand";
//            if (!(foundGridList is null) && (foundGridList.Count > 0))
//            {
//                int x = 1;
//                foreach (var foundGrid in foundGridList)
//                {
//                    var found = foundGrid.FoundInst;
//                    var foundLoc = found.Location as LocationPoint;
//                    var foundLocPoint = foundLoc.Point;
//                    var foundLx = found.Symbol.get_Parameter(BuiltInParameter.STRUCTURAL_FOUNDATION_LENGTH).AsDouble();
//                    var foundLy = found.Symbol.get_Parameter(BuiltInParameter.STRUCTURAL_FOUNDATION_WIDTH).AsDouble();
//                    //var foundTypeMark = found.Symbol.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString();
//                    //var foundZ = Math.Round(found.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble() * 0.3048, 3);
//                    var foundZ = MatchFoundationLevel(doc, found);
//                    List<Solid> foundSolids = found.GetAllSolidsAdvance(true);
//                    bool colIsVisible = false;
//                    double colX = 0.0;
//                    double colY = 0.0;
//                    //var forceCommand = "_set Point_" + x.ToString() + ".PointLoad.Fz";
                    
//                    var pointLoad = new FilteredElementCollector(doc).WhereElementIsNotElementType()
//                        .OfCategory(BuiltInCategory.OST_PointLoads).Where(e => e.get_Parameter(BuiltInParameter.LOAD_COMMENTS).AsString().Equals(found.Id.ToString())).FirstOrDefault();
//                    double pointLoadForce = 0.0;
//                    if (!(pointLoad is null))
//                    {
//                        pointLoadForce = Math.Round(pointLoad.get_Parameter(BuiltInParameter.LOAD_FORCE_FZ).AsDouble() * 0.0003048,3);
//                    }
//                    foreach (var foundSolid in foundSolids)
//                    {
//                        if (foundSolid.GraphicsStyleId.Equals(colStyleId))
//                        {
//                            colIsVisible = true;
//                            var colCentroid = foundSolid.ComputeCentroid();
//                            colX = Math.Round((colCentroid.X - projectPoint.Position.X) * 0.3048, 3);
//                            colY = Math.Round((colCentroid.Y - projectPoint.Position.Y) * 0.3048, 3);
//                            break;
//                        }
//                    }
//                    if (colIsVisible == false)
//                    {
                        
//                        if (found.Symbol.FamilyName == "基礎_通常")
//                        {
//                            var colOffsetX = found.Symbol.LookupParameter("柱型_移動_X").AsDouble();
//                            var colOffsetY = found.Symbol.LookupParameter("柱型_移動_Y").AsDouble();
//                            Transform transCol = Transform.CreateTranslation(found.GetTransform().BasisX * (-1) * (foundLx / 2 - colOffsetX))
//                                * Transform.CreateTranslation(found.GetTransform().BasisY * (-1) * (foundLy / 2 - colOffsetY));
//                            var transColPoint = transCol.OfPoint(foundLocPoint);
//                            colX = Math.Round(transColPoint.X * 0.3048, 3);
//                            colY = Math.Round(transColPoint.Y * 0.3048, 3);
//                        }
//                        else
//                        {
//                            colX = Math.Round(foundLocPoint.X * 0.3048, 3);
//                            colY = Math.Round(foundLocPoint.Y * 0.3048, 3);
//                        }
                        
//                    }
//                    var gridA = foundGrid.GridA.Name.CompareTo(foundGrid.GridB.Name) > 0 ? foundGrid.GridB.Name : foundGrid.GridA.Name;
//                    var gridB = foundGrid.GridA.Name.CompareTo(foundGrid.GridB.Name) < 0 ? foundGrid.GridB.Name : foundGrid.GridA.Name;
//                    var curStr = String.Format("{0};{1};{2};{3};{4};{5};{6};{7}", gridA, gridB
//                                , "_pointload", colX, colY, foundZ
//                                , tempForceCommand, pointLoadForce);
//                    colCommands.Add(curStr);
//                    XYZ pointLoadXYZ = new XYZ(colX / 0.3048, colY / 0.3048, foundZ / 0.3048);
//                    UserPointLoad userPointLoad = new UserPointLoad();
//                    userPointLoad.Point = pointLoadXYZ;
//                    userPointLoad.FoundId = found.Id;
//                    UserPointLoads.Add(userPointLoad);
//                    x++;
//                }
//            }
//            //colCommands.Sort((a, b) => a.CompareTo(b));
//            var sortColCommands = GridSorting(colCommands);
//            List<string> newColCommands = new List<string>();
//            if (!(sortColCommands is null) & (sortColCommands.Count > 0))
//            {
//                int x = 1;
//                foreach (var sortColCommand in sortColCommands)
//                {
//                    var forceCommand = "_set Point_" + x.ToString() + ".PointLoad.Fz";
//                    var newColCommand = sortColCommand.Replace(tempForceCommand, forceCommand);
//                    newColCommands.Add(newColCommand);
//                    x++;
//                }
//            }
//            /*
//            for (int j = 0; j < colCommands.Count; j++)
//            {
//                var forceCommand = "_set Point_" + (j + 1).ToString() + ".PointLoad.Fz";
//                colCommands[j] = String.Format("{0};{1}", colCommands[j], forceCommand);
//            }
//            */
//            string colDataName = "柱荷重";
//            newColCommands.Insert(0, colDataName);
//            return (newColCommands, UserPointLoads);
//        }
//        public static List<string> GridSorting(List<string> list)
//        {
//            if (list == null || list.Count == 0) return null;
//            string regex = @"(?<label>[A-z]*)(?<number>\d*)(?<postfix>.*)";
//            var sortList = list.OrderBy(x => Regex.Match(x.Split(';').ElementAt(0), regex).Groups["label"].Value)
//                               .ThenBy(x =>
//                               {
//                                   var number = Regex.Match(x.Split(';').ElementAt(0), regex)?.Groups["number"].Value;
//                                   return string.IsNullOrWhiteSpace(number) ? int.MinValue : int.Parse(number);
//                               })
//                               .ThenBy(x => Regex.Match(x.Split(';').ElementAt(0), regex).Groups["postfix"].Value)
//                               .ThenBy(y => Regex.Match(y.Split(';').ElementAt(1), regex).Groups["label"].Value)
//                               .ThenBy(y => {
//                                   var number = Regex.Match(y.Split(';').ElementAt(1), regex)?.Groups["number"].Value;
//                                   return string.IsNullOrWhiteSpace(number) ? int.MinValue : int.Parse(number);
//                               })
//                               .ThenBy(y => Regex.Match(y.Split(';').ElementAt(1), regex).Groups["postfix"].Value);
//            return sortList.ToList();
//        }

//    }
//}
