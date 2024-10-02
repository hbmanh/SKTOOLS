using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;

namespace SKRevitAddins.Utils
{
    public static class TSolidUtils
    {
        public static PlanarFace GetMaxFace(this Solid solid)
        {
            PlanarFace maxFace = null;
            for (int i = 0; i < solid.Faces.Size; i++)
            {
                var currentFace = solid.Faces.get_Item(i);
                if (currentFace is PlanarFace)
                {
                    if (maxFace.Area < currentFace.Area)
                    {
                        maxFace = currentFace as PlanarFace;
                    }
                }
            }
            return maxFace;
        }
        public static PlanarFace GetBottomPlanarFace(this Solid solid)
        {
            if ((solid == null) || (solid.Volume == 0))
            {
                return null;
            }
            PlanarFace bottomFace = null;
            for (int i = 0; i < solid.Faces.Size; i++)
            {
                var currentFace = solid.Faces.get_Item(i);
                if (currentFace is PlanarFace)
                {
                    if (Math.Round((currentFace as PlanarFace).FaceNormal.Z,1) == -1.0)
                    {
                        bottomFace = currentFace as PlanarFace;
                    }
                }
            }
            return bottomFace;
        }
        public static PlanarFace GetTopPlanarFace(this Solid solid)
        {
            if ((solid == null) || (solid.Volume == 0))
            {
                return null;
            }
            PlanarFace topFace = null;
            for (int i = 0; i < solid.Faces.Size; i++)
            {
                var currentFace = solid.Faces.get_Item(i);
                if (currentFace is PlanarFace)
                {
                    if (Math.Round((currentFace as PlanarFace).FaceNormal.Z,1) == 1.0)
                    {
                        topFace = currentFace as PlanarFace;
                    }
                }
            }
            return topFace;
        }
        public static List<PlanarFace> GetTopPlanarFaceList(this Solid solid, bool includeSlope = false)
        {
            List<PlanarFace> topFaceList = new List<PlanarFace>();
            for (int i = 0; i < solid.Faces.Size; i++)
            {
                var currentFace = solid.Faces.get_Item(i);
                if (currentFace is PlanarFace)
                {
                    if (includeSlope == false)
                    {
                        if (Math.Round((currentFace as PlanarFace).FaceNormal.Z, 1) == 1.0)
                        {
                            topFaceList.Add(currentFace as PlanarFace);
                        }
                    }
                    else
                    {
                        if ((currentFace as PlanarFace).FaceNormal.AngleTo(XYZ.BasisZ) < Math.PI / 2)
                        {
                            topFaceList.Add(currentFace as PlanarFace);
                        }
                    }
                }
            }
            return topFaceList;
        }
        public static List<PlanarFace> GetBottomPlanarFaceList(this Solid solid, bool includeSlope = false)
        {
            List<PlanarFace> botFaceList = new List<PlanarFace>();
            for (int i = 0; i < solid.Faces.Size; i++)
            {
                var currentFace = solid.Faces.get_Item(i);
                if (currentFace is PlanarFace)
                {
                    if (includeSlope == false)
                    {
                        if (Math.Round((currentFace as PlanarFace).FaceNormal.Z, 1) == -1.0)
                        {
                            botFaceList.Add(currentFace as PlanarFace);
                        }
                    }
                    else
                    {
                        if ((currentFace as PlanarFace).FaceNormal.AngleTo(XYZ.BasisZ * -1) < Math.PI / 2)
                        {
                            botFaceList.Add(currentFace as PlanarFace);
                        }
                    }
                }
            }
            return botFaceList;
        }
        public static XYZ GetPlanarFaceCentroid(PlanarFace planarFace)
        {
            var edgeArray = planarFace.EdgeLoops.get_Item(0);
            List<Edge> edges = new List<Edge>();
            for (int i = 0; i < edgeArray.Size; i++)
            {
                var currentEdge = edgeArray.get_Item(i);
                if (currentEdge.Visibility == Visibility.Visible)
                {
                    edges.Add(currentEdge);
                }
            }
            var midPoint = default(XYZ);
            if (edges.Count == 4)
            {
                double sumX = 0, sumY = 0, z = 0;
                foreach (var edge in edges)
                {
                    sumX += edge.Tessellate().ElementAt(0).X + edge.Tessellate().ElementAt(1).X;
                    sumY += edge.Tessellate().ElementAt(0).Y + edge.Tessellate().ElementAt(1).Y;
                    z = edge.Tessellate().ElementAt(0).Z;
                }
                midPoint = new XYZ(sumX / 8, sumY / 8, z);
            }
            return midPoint;
        }
        public static XYZ GetMiddlePoint(XYZ point1, XYZ point2)
        {
            return new XYZ((point1.X + point2.X) / 2, (point1.Y + point2.Y) / 2, (point1.Z + point2.Z) / 2);
        }
        public static bool CutGeometry(Document doc, Element beCutInstance, Element cutInstance)
        {
            try
            {
                var beCutSolids = beCutInstance.GetAllSolids(true);
                var cutSolids = cutInstance.GetAllSolids(true);
                bool result = false;
                if (beCutSolids != null &&
                    beCutSolids.Count != 0 &&
                    cutSolids != null &&
                    cutSolids.Count != 0)
                {
                    var intersection = BooleanOperationsUtils.ExecuteBooleanOperation(beCutSolids.First(), cutSolids.First(), BooleanOperationsType.Intersect);
                    if (intersection.Volume != 0)
                    {
                        SolidSolidCutUtils.AddCutBetweenSolids(doc, beCutInstance, cutInstance);
                        result = true;
                    }
                }
                return result;
            }
            
            catch (Exception e)
            {

                throw e;
            }
        
        }
        public static Solid CreateTrapezoid(XYZ botP1, XYZ botP2, XYZ botP3, XYZ botP4, XYZ dirX, XYZ dirY, double trapSize)
        {
            // Return Null if One Input is Null
            if ((botP1 is null) || (botP2 is null) || (botP3 is null)
                || (botP4 is null) || (dirX is null) || (dirY is null))
            {
                return null;
            }

            // Based on Bottom Points, Create Bottom Curves
            Curve botC1 = Line.CreateBound(botP1, botP2);
            Curve botC2 = Line.CreateBound(botP2, botP3);
            Curve botC3 = Line.CreateBound(botP3, botP4);
            Curve botC4 = Line.CreateBound(botP4, botP1);

            // Based on Bottom Curves, Create A Bottom CurveLoop
            List<Curve> botCs = new List<Curve>();
            botCs.Add(botC1);
            botCs.Add(botC2);
            botCs.Add(botC3);
            botCs.Add(botC4);
            CurveLoop botCl = CurveLoop.Create(botCs);

            // Create Transforms
            Transform tf1 = Transform.CreateTranslation(dirX * trapSize)
                * Transform.CreateTranslation(dirY * trapSize)
                * Transform.CreateTranslation(XYZ.BasisZ * trapSize);
            Transform tf2 = Transform.CreateTranslation(dirX * trapSize)
                * Transform.CreateTranslation(dirY * -trapSize)
                * Transform.CreateTranslation(XYZ.BasisZ * trapSize);
            Transform tf3 = Transform.CreateTranslation(dirX * -trapSize)
                * Transform.CreateTranslation(dirY * -trapSize)
                * Transform.CreateTranslation(XYZ.BasisZ * trapSize);
            Transform tf4 = Transform.CreateTranslation(dirX * -trapSize)
                * Transform.CreateTranslation(dirY * trapSize)
                * Transform.CreateTranslation(XYZ.BasisZ * trapSize);

            // Based on Transforms, Create Top Points
            XYZ topP1 = tf1.OfPoint(botP1);
            XYZ topP2 = tf2.OfPoint(botP2);
            XYZ topP3 = tf3.OfPoint(botP3);
            XYZ topP4 = tf4.OfPoint(botP4);

            // Based on Top Points, Create Top Curves
            Curve topC1 = Line.CreateBound(topP1, topP2);
            Curve topC2 = Line.CreateBound(topP2, topP3);
            Curve topC3 = Line.CreateBound(topP3, topP4);
            Curve topC4 = Line.CreateBound(topP4, topP1);

            // Based on Top Curves, Create A Top CurveLoop
            List<Curve> topCs = new List<Curve>();
            topCs.Add(topC1);
            topCs.Add(topC2);
            topCs.Add(topC3);
            topCs.Add(topC4);
            CurveLoop topCl = CurveLoop.Create(topCs);

            Solid trapSolid = GeometryCreationUtilities.CreateBlendGeometry(botCl, topCl, null);
            return trapSolid;
        }
        public static Solid GetFloorOpeningSolid (this Floor floor, Document doc)
        {
            Solid openingExtrudeUnion = null;
            var openingColl = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_FloorOpening)
                .OfClass(typeof(Opening))
                .Cast<Opening>()
                .Where(e => e.Host.Id.Equals(floor.Id));
            var floorGlOffset = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble();
            Transform tfFloorGlOffset = Transform.CreateTranslation(XYZ.BasisZ * floorGlOffset);
            var floorThickness = floor.FloorType.get_Parameter(BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM).AsDouble();
            if (openingColl.Count() > 0)
            {
                foreach (var opening in openingColl)
                {
                    CurveArray openingCa = opening.BoundaryCurves;
                    CurveLoop openingCl = new CurveLoop();
                    openingCl.AppendFromCurveArray(openingCa);
                    openingCl.Transform(tfFloorGlOffset);
                    List<CurveLoop> openingCls = new List<CurveLoop>();
                    openingCls.Add(openingCl);
                    Solid openingExtrude = null;
                    try
                    {
                        openingExtrude = GeometryCreationUtilities.CreateExtrusionGeometry(openingCls, XYZ.BasisZ * -1, floorThickness);
                        // openingExtrude.BakeSolidToDirectShape(doc);
                    }
                    catch
                    {
                        Debug.WriteLine("Can't create opening extrusion");
                    }

                    if (openingExtrude == null) return null;
                    
                    if (openingExtrudeUnion is null)
                    {
                        openingExtrudeUnion = openingExtrude;
                    }
                    else
                    {
                        openingExtrudeUnion = BooleanOperationsUtils.ExecuteBooleanOperation(openingExtrudeUnion, openingExtrude, BooleanOperationsType.Union);
                    }
                }
            }
            return openingExtrudeUnion;
        }
        public static Solid FloorBotFaceExtrude(this Floor floor, Document doc, XYZ dir, double dist, bool ignoreOpening = true)
        {
            if (floor is null)
            {
                return null;
            }
            Solid floorBotFaceExtrude = null;
            Solid floorSolidUnion = null;
            var floorOpeningSolid = floor.GetFloorOpeningSolid(doc);
            // floorOpeningSolid.BakeSolidToDirectShape(doc);
            var floorSolids = floor.GetAllSolidsAdvance(true)
                .Where(s => (s != null) && (s.Volume > 0)).ToList();
            if (floorSolids.Count > 0)
            {
                foreach (var floorSolid in floorSolids)
                {
                    if (floorSolidUnion is null)
                    {
                        floorSolidUnion = floorSolid;
                    }
                    else
                    {
                        floorSolidUnion = BooleanOperationsUtils.ExecuteBooleanOperation(floorSolidUnion, floorSolid, BooleanOperationsType.Union);
                    }
                }
            }
            if (ignoreOpening == true)
            {
                if ((floorOpeningSolid != null) && (floorOpeningSolid.Volume > 0))
                {
                    floorSolidUnion = BooleanOperationsUtils.ExecuteBooleanOperation(floorSolidUnion, floorOpeningSolid, BooleanOperationsType.Union);
                }
            }

            var floorBotFaces = floorSolidUnion.GetBottomPlanarFaceList(true);
            foreach (var floorBotFace in floorBotFaces)
            {
                var floorBotFaceCurveLoops = floorBotFace.GetEdgesAsCurveLoops();
                var tempFloorBotFaceExtrude = GeometryCreationUtilities.CreateExtrusionGeometry(
                            floorBotFaceCurveLoops, dir, dist);

                if (floorBotFaceExtrude is null)
                {
                    floorBotFaceExtrude = tempFloorBotFaceExtrude;
                }
                else
                {
                    floorBotFaceExtrude = BooleanOperationsUtils.ExecuteBooleanOperation(
                        floorBotFaceExtrude, tempFloorBotFaceExtrude, BooleanOperationsType.Union);
                }

                if (floorBotFaceCurveLoops.Count > 0)
                {
                    /*foreach (var floorBotFaceCurveLoop in floorBotFaceCurveLoops)
                    {
                        List<CurveLoop> floorBotFaceCurveLoopList = new List<CurveLoop>();
                        floorBotFaceCurveLoopList.Add(floorBotFaceCurveLoop);
                        var tempFloorBotFaceExtrude = GeometryCreationUtilities.CreateExtrusionGeometry(
                            floorBotFaceCurveLoopList, dir, dist);
                        if (floorBotFaceExtrude is null)
                        {
                            floorBotFaceExtrude = tempFloorBotFaceExtrude;
                        }
                        else
                        {
                            floorBotFaceExtrude = BooleanOperationsUtils.ExecuteBooleanOperation(
                                floorBotFaceExtrude, tempFloorBotFaceExtrude, BooleanOperationsType.Union);
                        }
                    }*/
                }
            }
            using (Transaction tx = new Transaction(doc))
            {
                /*tx.Start("Bake");
                floorBotFaceExtrude.BakeSolidToDirectShape(doc);
                tx.Commit();*/
            }
            return floorBotFaceExtrude;
        }
      
        public static Solid FloorBotFaceExtrudeUnion (this List<Floor> floors, Document doc, XYZ dir, double dist, bool ignoreOpening = true)
        {
            if (floors is null)
            {
                return null;
            }
            Solid floorBotFaceExtrudeUnion = null;
            foreach (var floor in floors)
            {
                Solid floorBotFaceExtrude = floor.FloorBotFaceExtrude(doc, dir, dist, ignoreOpening);
                if (floorBotFaceExtrudeUnion is null)
                {
                    floorBotFaceExtrudeUnion = floorBotFaceExtrude;
                }
                else
                {
                    floorBotFaceExtrudeUnion = BooleanOperationsUtils.ExecuteBooleanOperation(
                        floorBotFaceExtrudeUnion, floorBotFaceExtrude, BooleanOperationsType.Union);
                }
            }            
            return floorBotFaceExtrudeUnion;
        }
        public static Solid FoundMainAndYoboriSolid(this FamilyInstance found)
        {
            if ((found is null) || (found.Category.Id.IntegerValue != -2001300))
            {
                return null;
            }
            var foundMainSolid = found.GetAllSolidsAdvance(true)
                .Where(s => (s != null) && (s.Volume > 0)
                && (s.GraphicsStyleId.IntegerValue.Equals(625783))).First(); // Main Solid
            var foundYoboriSolid = found.GetAllSolidsAdvance(true)
                .Where(s => (s != null) && (s.Volume > 0)
                && (s.GraphicsStyleId.IntegerValue.Equals(2185417))).First(); // Yobori Solid
            Solid foundMainAndYoboriSolid = BooleanOperationsUtils.ExecuteBooleanOperation(foundMainSolid, foundYoboriSolid, BooleanOperationsType.Union);
            return foundMainAndYoboriSolid;
        }
        public static Solid FwfgMainAndYoboriSolid(this FamilyInstance fwfg)
        {
            if ((fwfg is null) || (fwfg.Category.Id.IntegerValue != -2001320))
            {
                return null;
            }
            var fwfgMainSolids = fwfg.GetAllSolidsAdvance(true)
                .Where(s => (s != null) && (s.Volume > 0)
                && (s.GraphicsStyleId.IntegerValue.Equals(2192344)));
            Solid fwfgMainSolid = null;
            if (fwfgMainSolids.Count() > 0)
            {
                fwfgMainSolid = fwfg.GetAllSolidsAdvance(true)
                .Where(s => (s != null) && (s.Volume > 0)
                && (s.GraphicsStyleId.IntegerValue.Equals(2192344))).First(); // Main Solid
            }
            var fwfgYoboriSolids = fwfg.GetAllSolidsAdvance(true)
                .Where(s => (s != null) && (s.Volume > 0)
                && (s.GraphicsStyleId.IntegerValue.Equals(3437393)));
            Solid fwfgYoboriSolid = null;
            if (fwfgYoboriSolids.Count() > 0)
            {
                fwfgYoboriSolid = fwfg.GetAllSolidsAdvance(true)
                .Where(s => (s != null) && (s.Volume > 0)
                && (s.GraphicsStyleId.IntegerValue.Equals(3437393))).First(); // Yobori Solid
            }
            Solid fwfgMainAndYoboriSolid = null;
            if ((fwfgMainSolid != null) && (fwfgMainSolid.Volume > 0)
                && (fwfgYoboriSolid != null) && (fwfgYoboriSolid.Volume > 0))
            {
                fwfgMainAndYoboriSolid = BooleanOperationsUtils.ExecuteBooleanOperation(fwfgMainSolid, fwfgYoboriSolid, BooleanOperationsType.Union);
            }
            /*
            else
            {
                var h = fwfg.Symbol.LookupParameter("h").AsDouble();
                fwfgMainAndYoboriSolid = fwfg.FwfgTopFaceExtrudeByLocLine(XYZ.BasisZ * -1, h);
            }*/
            return fwfgMainAndYoboriSolid;
        }
        public static Solid FoundTopFaceExtrude(this FamilyInstance found, XYZ dir, double extrudeDist)
        {
            extrudeDist = (extrudeDist == 0) ? 30 : extrudeDist;
            var foundMainAndYoboriSolid = found.FoundMainAndYoboriSolid();
            var foundTopFace = foundMainAndYoboriSolid.GetTopPlanarFace();
            var foundTopFaceCurveLoops = foundTopFace.GetEdgesAsCurveLoops();
            Solid foundTopFaceCurveLoopExtrudeUp = GeometryCreationUtilities.CreateExtrusionGeometry(
                foundTopFaceCurveLoops, dir, extrudeDist);
            return foundTopFaceCurveLoopExtrudeUp;
        }
        public static Solid JibanExtrude(this FamilyInstance jiban, XYZ dir, double dist, bool transformToGl = false)
        {
            if (jiban is null || jiban.GetAllSolidsAdvance(true).Count == 0)
            {
                return null;
            }


            // Register Variables
            var jibanThickness = jiban.Symbol.get_Parameter(BuiltInParameter.CURTAIN_WALL_SYSPANEL_THICKNESS).AsDouble();
            var jibanHost = jiban.Host as FootPrintRoof;
            var jibanSolid = jiban.GetAllSolidsAdvance(true);
            var jibanCls = jibanSolid
                                   .UnionSolidList()
                                   .GetBottomPlanarFace()
                                   .GetEdgesAsCurveLoops();
            var jibanGlOffset = jibanHost.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM).AsDouble();
            jibanGlOffset = (transformToGl == false) ? 1 : jibanGlOffset;
            Transform tfJibanBotToTop = Transform.CreateTranslation(XYZ.BasisZ * jibanThickness * jibanGlOffset);
            List<CurveLoop> jibanClsTf = new List<CurveLoop>();
            foreach (var jibanCl in jibanCls)
            {
                jibanCl.Transform(tfJibanBotToTop);
                jibanClsTf.Add(jibanCl);
            }
            Solid jibanExtrude = GeometryCreationUtilities.CreateExtrusionGeometry(jibanClsTf, dir, dist);
            return jibanExtrude;
        }
        public static Solid JibanExtrudeUnion(this List<FamilyInstance> jibans, XYZ dir, double dist, bool transformToGl = false)
        {
            if ((jibans is null) || (jibans.Count == 0))
            {
                return null;
            }
            Solid jibanExtrudeUnion = null;
            foreach (var jiban in jibans)
            {
                Solid jibanExtrude = jiban.JibanExtrude(dir, dist, transformToGl);
                if (jibanExtrude == null) continue;
                if ((jibanExtrudeUnion is null) || (jibanExtrudeUnion.Volume == 0))
                {
                    jibanExtrudeUnion = jibanExtrude;
                }
                else
                {
                    jibanExtrudeUnion = BooleanOperationsUtils.ExecuteBooleanOperation(
                        jibanExtrudeUnion, jibanExtrude, BooleanOperationsType.Union);
                }
            }
            return jibanExtrudeUnion;
        }
        public static Solid FwfgTopFaceExtrudeByLocLine(this FamilyInstance fwfg, XYZ dir, double dist)
        {
            if (fwfg is null)
            {
                return null;
            }
            var fwfgCurve = (fwfg.Location as LocationCurve).Curve;
            var b = fwfg.Symbol.LookupParameter("b").AsDouble();
            var yobori = 0.1;
            var width = b + yobori;
            var rightCurve = fwfgCurve.CreateOffset(width / 2, XYZ.BasisZ);
            var leftCurve = fwfgCurve.CreateOffset(-width / 2, XYZ.BasisZ);
            var rp1 = rightCurve.GetEndPoint(0);
            var rp2 = rightCurve.GetEndPoint(1);
            var lp1 = leftCurve.GetEndPoint(0);
            var lp2 = leftCurve.GetEndPoint(1);

            var l1 = rightCurve;
            var l2 = Line.CreateBound(rp2, lp2);
            var l3 = Line.CreateBound(lp2, lp1);
            var l4 = Line.CreateBound(lp1, rp1);

            CurveLoop curveLoop = new CurveLoop();
            curveLoop.Append(l1);
            curveLoop.Append(l2);
            curveLoop.Append(l3);
            curveLoop.Append(l4);
            List<CurveLoop> curveLoopList = new List<CurveLoop>();
            curveLoopList.Add(curveLoop);
            Solid extrude = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoopList, dir, dist);
            return extrude;
        }
        public static Solid FwfgTopFaceExtrude(this FamilyInstance fwfg, XYZ dir, double dist)
        {
            if (fwfg is null)
            {
                return null;
            }
            var fwfgMainSolids = fwfg.GetAllSolidsAdvance(true)
                .Where(s => (s != null) && (s.Volume > 0) && (s.GraphicsStyleId.IntegerValue.Equals(2192344)));
            Solid fwfgMainSolid = null;
            if (fwfgMainSolids.Count() > 0)
            {
                fwfgMainSolid = fwfgMainSolids.First();
            }
            var fwfgYoboriSolids = fwfg.GetAllSolidsAdvance(true)
                .Where(s => (s != null) && (s.Volume > 0) && (s.GraphicsStyleId.IntegerValue.Equals(3437393)));
            Solid fwfgYoboriSolid = null;
            if (fwfgYoboriSolids.Count() > 0)
            {
                fwfgYoboriSolid = fwfgYoboriSolids.First();
            }
            Solid fwfgTopFaceCurveLoopExtrude = null;
            if ((fwfgMainSolid != null) && (fwfgMainSolid.Volume > 0) 
                && (fwfgYoboriSolid != null) && (fwfgYoboriSolid.Volume > 0))
            {
                var fwfgMainAndYoboriSolid = BooleanOperationsUtils.ExecuteBooleanOperation(fwfgMainSolid, fwfgYoboriSolid, BooleanOperationsType.Union);
                var fwfgTopFace = fwfgMainAndYoboriSolid.GetTopPlanarFace();
                var fwfgFaceCurveLoops = fwfgTopFace.GetEdgesAsCurveLoops();
                fwfgTopFaceCurveLoopExtrude = GeometryCreationUtilities.CreateExtrusionGeometry(
                fwfgFaceCurveLoops, dir, dist);
            }
            else
            {
                fwfgTopFaceCurveLoopExtrude = fwfg.FwfgTopFaceExtrudeByLocLine(dir, dist);
            }
            return fwfgTopFaceCurveLoopExtrude;
        }
        public static Solid UnionInstanceSolids(this Element ele)
        {
            if (ele is null)
            {
                return null;
            }
            var instanceSolids = ele.GetAllSolidsAdvance(true).ToList();
            Solid instanceSolidUnion = null;
            foreach (var instanceSolid in instanceSolids)
            {
                if ((instanceSolid != null) && (instanceSolid.Volume > 0))
                {
                    if (instanceSolidUnion is null)
                    {
                        instanceSolidUnion = instanceSolid;
                    }
                    else
                    {
                        instanceSolidUnion = BooleanOperationsUtils.ExecuteBooleanOperation(instanceSolidUnion, instanceSolid, BooleanOperationsType.Union);
                    }
                }
            }
            return instanceSolidUnion;
        }
        public static Solid UnionInstColl(this List<Element> listEle)
        {
            if ((listEle is null) && (listEle.Count == 0))
            {
                return null;
            }
            Solid instanceCollUnion = null;
            foreach (var instance in listEle)
            {
                var instanceSolidUnion = instance.UnionInstanceSolids();
                if ((instanceSolidUnion != null) && (instanceSolidUnion.Volume > 0))
                {
                    if (instanceCollUnion is null)
                    {
                        instanceCollUnion = instanceSolidUnion;
                    }
                    else
                    {
                        instanceCollUnion = BooleanOperationsUtils.ExecuteBooleanOperation(
                            instanceCollUnion, instanceSolidUnion, BooleanOperationsType.Union);
                    }
                }
            }
            return instanceCollUnion;
        }
        public static List<Solid> FoundNearFwfg(this FamilyInstance fwfg, List<FamilyInstance> founds)
        {
            if ((fwfg is null) || (founds is null))
            {
                return null;
            }
            var fwfgSolidUnion = fwfg.GetAllSolidsAdvance(true).Where(s => (s != null) && (s.Volume > 0)
                && s.GraphicsStyleId.IntegerValue.Equals(2192344)).First();
            var fwfgFaces = fwfgSolidUnion.Faces;
            List<PlanarFace> fwfgVerFaces = new List<PlanarFace>();
            foreach (PlanarFace fwfgFace in fwfgFaces)
            {
                if (fwfgFace.FaceNormal.Z.Equals(0))
                {
                    fwfgVerFaces.Add(fwfgFace);
                }
            }
            fwfgVerFaces.Sort((a, b) => a.Area.CompareTo(b.Area));
            var fwfgSideFace1 = fwfgVerFaces[0];
            var fwfgSideFace2 = fwfgVerFaces[1];

            var fwfgSideCurveLoop1 = fwfgSideFace1.GetEdgesAsCurveLoops().First();
            var fwfgSideCurveLoop2 = fwfgSideFace2.GetEdgesAsCurveLoops().First();

            CurveLoop offsetCl1 = CurveLoop.CreateViaOffset(fwfgSideCurveLoop1, 2, fwfgSideFace1.FaceNormal);
            CurveLoop offsetCl2 = CurveLoop.CreateViaOffset(fwfgSideCurveLoop2, 2, fwfgSideFace2.FaceNormal);

            List<CurveLoop> offsetCl1List = new List<CurveLoop>();
            List<CurveLoop> offsetCl2List = new List<CurveLoop>();

            offsetCl1List.Add(offsetCl1);
            offsetCl2List.Add(offsetCl2);

            Solid fwfgSideExtrude1 = GeometryCreationUtilities.CreateExtrusionGeometry(offsetCl1List, fwfgSideFace1.FaceNormal, 5);
            Solid fwfgSideExtrude2 = GeometryCreationUtilities.CreateExtrusionGeometry(offsetCl2List, fwfgSideFace2.FaceNormal, 5);

            List<Solid> fwfgSideSolids = new List<Solid>();
            fwfgSideSolids.Add(fwfgSideExtrude1);
            fwfgSideSolids.Add(fwfgSideExtrude2);
            return fwfgSideSolids;
        }
        public static List<PlanarFace> GetTwoSideFaces(this Solid solid)
        {
            List<PlanarFace> verFaces = new List<PlanarFace>();
            var solidFaces = solid.Faces;
            foreach (PlanarFace solidFace in solidFaces)
            {
                if (Math.Round(solidFace.FaceNormal.Z,3).Equals(0))
                {
                    verFaces.Add(solidFace);
                }
            }
            verFaces.Sort((a, b) => a.Area.CompareTo(b.Area));

            List<PlanarFace> twoSideFaces = new List<PlanarFace>();
            if (verFaces.Count >= 2)
            {
                twoSideFaces.Add(verFaces[0]);
                twoSideFaces.Add(verFaces[1]);
            }
            
            return twoSideFaces;
        }
        public static Line GetBottomLineOfFace(this Face face)
        {
            if (face is null || face.Area  <= 0)
            {
                return null;
            }
            var curveLoop = face.GetEdgesAsCurveLoops().First();
            Line botLine = null;
            XYZ botLineOrigin = new XYZ(double.MaxValue, double.MaxValue, double.MaxValue);
            foreach (Line line in curveLoop)
            {
                var lineDir = line.Direction;
                var lineOrigin = line.Origin;
                if (Math.Round(lineDir.DotProduct(XYZ.BasisZ),3) == 0
                    && lineOrigin.Z < botLineOrigin.Z)
                {
                    botLine = line;
                    botLineOrigin = lineOrigin;
                }
            }
            return botLine;
        }
        public static List<CurveLoopWithNormal> GetTwoSideCurveLoops (this Solid solid)
        {
            var twoSideFaces = solid.GetTwoSideFaces();
            List<CurveLoopWithNormal> twoSideFaceCurveLoops = new List<CurveLoopWithNormal>();
            foreach (var sideFace in twoSideFaces)
            {
                var curveLoop = sideFace.GetEdgesAsCurveLoops();
                CurveLoopWithNormal curveLoopWithNormal = new CurveLoopWithNormal();
                curveLoopWithNormal.CurveLoops = curveLoop;
                curveLoopWithNormal.Normal = sideFace.FaceNormal;
                twoSideFaceCurveLoops.Add(curveLoopWithNormal);
            }
            return twoSideFaceCurveLoops;
        }
        public static List<Solid> ExtrudeTwoSideFaces (this Solid solid)
        {
            List<Solid> twoSideFaceExtrudes = new List<Solid>();
            var twoSideFaceCurveLoops = solid.GetTwoSideCurveLoops();
            foreach (var sideFaceCurveLoop in twoSideFaceCurveLoops)
            {
                try
                {
                    Solid sideFaceExtrude = GeometryCreationUtilities.CreateExtrusionGeometry(
                    sideFaceCurveLoop.CurveLoops, sideFaceCurveLoop.Normal, 6);

                    twoSideFaceExtrudes.Add(sideFaceExtrude);
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {

                }
            }
            return twoSideFaceExtrudes;
        }
        public class CurveLoopWithNormal
        {
            public IList<CurveLoop> CurveLoops { get; set; }
            public XYZ Normal { get; set; }
        }
        public static Solid GetFoundGataSolid (this FamilyInstance found)
        {
            if (found is null)
            {
                return null;
            }
            var foundGataSolid = found.GetAllSolidsAdvance(true)
                .Where(s => (s != null) && (s.Volume > 0)
                && s.GraphicsStyleId.IntegerValue.Equals(868429)).First();
            return foundGataSolid;
        }
        public static Solid GetDifferenceIfIntersect(this Solid solid1, Solid solid2)
        {
            Solid diffSolid = null;
            try
            {
                var intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(
                                        solid1, solid2, BooleanOperationsType.Intersect);
                if ((intersectSolid != null) && (intersectSolid.Volume > 0))
                {

                    diffSolid = BooleanOperationsUtils.ExecuteBooleanOperation(
                       solid1, solid2, BooleanOperationsType.Difference);
                }
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {

            }
            return diffSolid;
        }
        public static List<XYZ> GetRecPointsFromCenter(this XYZ centerPoint, XYZ dirX, XYZ dirY, double xOffset, double yOffset)
        {
            Transform tf1 = Transform.CreateTranslation(dirX * xOffset) * Transform.CreateTranslation(dirY * yOffset);
            Transform tf2 = Transform.CreateTranslation(dirX * xOffset) * Transform.CreateTranslation(dirY * yOffset * -1);
            Transform tf3 = Transform.CreateTranslation(dirX * xOffset * -1) * Transform.CreateTranslation(dirY * yOffset * -1);
            Transform tf4 = Transform.CreateTranslation(dirX * xOffset * -1) * Transform.CreateTranslation(dirY * yOffset);

            XYZ p1 = tf1.OfPoint(centerPoint);
            XYZ p2 = tf2.OfPoint(centerPoint);
            XYZ p3 = tf3.OfPoint(centerPoint);
            XYZ p4 = tf4.OfPoint(centerPoint);

            List<XYZ> recPoints = new List<XYZ>();

            recPoints.Add(p1);
            recPoints.Add(p2);
            recPoints.Add(p3);
            recPoints.Add(p4);

            return recPoints;
        }
        public static DirectShape BakeSolidToDirectShape(this Solid solid, Document doc, ElementId categoryId = null, string comments = "", string hostTypeMark = "", ElementId hostId = null)
        {
            DirectShape ds = null;
            if (categoryId == null)
            {
                categoryId = new ElementId(BuiltInCategory.OST_GenericModel);
            }
            if (hostId == null)
            {
                hostId = new ElementId(BuiltInCategory.OST_GenericModel);
            }
            if ((solid != null) && (solid.Volume > 0))
            {
                ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                ds.ApplicationId = "Application id";
                ds.ApplicationDataId = "Geometry object id";
                ds.SetShape(new GeometryObject[] { solid });
                ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(comments);
                //ds.LookupParameter("ホスト_Id").Set(hostId.ToString());
                //ds.LookupParameter("ホスト").Set(hostTypeMark);
                //ds.LookupParameter("体積").Set(solid.Volume);
                //if (solid.CheckIfPrism().Item1)
                //{
                //    ds.LookupParameter("平均厚").Set(solid.CheckIfPrism().Item3);
                //    ds.LookupParameter("面積").Set(solid.CheckIfPrism().Item4);
                //}
                
            }
            return ds;
        }
        public static CurveLoop TransformCurveLoopToGl(this CurveLoop curveLoop)
        {
            List<Curve> newCurveList = new List<Curve>();
            foreach (var curve in curveLoop)
            {
                var op1 = curve.GetEndPoint(0);
                var op2 = curve.GetEndPoint(1);
                var dist1 = op1.Z;
                var dist2 = op2.Z;
                // This Utils works for Slope Line but does not works well for Sloped Arcs
                if (dist1 != dist2)
                {
                    var np1 = new XYZ(op1.X, op1.Y, 0);
                    var np2 = new XYZ(op2.X, op2.Y, 0);
                    Curve newCurve = Line.CreateBound(np1, np2);
                    newCurveList.Add(newCurve);
                }
                else
                {
                    Transform tf = Transform.CreateTranslation(XYZ.BasisZ * dist1 * -1);
                    Curve newCurve = curve.CreateTransformed(tf);
                    newCurveList.Add(newCurve);
                }
            }
            CurveLoop glCurveLoop = CurveLoop.Create(newCurveList);
            return glCurveLoop;
        }

        public static Face GetFlatFaceOfSolid (this Solid solid)
        {
            if ((solid is null) || (solid.Volume == 0))
            {
                return null;
            }
            XYZ centroid = solid.ComputeCentroid();
            Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, centroid);
            Curve halfCircle1 = Arc.Create(plane, 1000, 0, 180 * Math.PI / 180) as Curve;
            Curve halfCircle2 = Arc.Create(plane, 1000, 180 * Math.PI / 180, 360 * Math.PI / 180) as Curve;
            CurveLoop curveLoop = new CurveLoop();
            curveLoop.Append(halfCircle1);
            curveLoop.Append(halfCircle2);
            IList<CurveLoop> curveLoopList = new List<CurveLoop>();
            curveLoopList.Add(curveLoop);
            Solid newSolid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoopList, XYZ.BasisZ * -1, 50);
            Solid intersect = BooleanOperationsUtils.ExecuteBooleanOperation(solid, newSolid, BooleanOperationsType.Intersect);
            PlanarFace topFace = intersect.GetTopPlanarFace();
            return topFace;
        }

        public static void UnJoinElements(this List<Element> eles, Document doc)
        {
            foreach (var ele1 in eles)
            {
                foreach (var ele2 in eles)
                {
                    if (ele1.Id != ele2.Id)
                    {
                        bool isJoined = JoinGeometryUtils.AreElementsJoined(doc, ele1, ele2);
                        if (isJoined is true)
                        {
                            JoinGeometryUtils.UnjoinGeometry(doc, ele1, ele2);
                        }
                    }
                }
            }
        }
        public static List<Element> UnJoinOneWithOthers (this Element ele, List<Element> otherEles, Document doc)
        {
            List<Element> unjoinEleList = new List<Element>();
            unjoinEleList.Add(ele);
            if (otherEles.Count > 0)
            {
                foreach (var otherEle in otherEles)
                {
                    if (otherEle.Id != ele.Id)
                    {
                        bool isJoined = JoinGeometryUtils.AreElementsJoined(doc, ele, otherEle);
                        if (isJoined is true)
                        {
                            JoinGeometryUtils.UnjoinGeometry(doc, ele, otherEle);
                            if (!(unjoinEleList.Contains(otherEle)))
                            {
                                unjoinEleList.Add(otherEle);
                            }
                        }
                    }
                }
            }
            return unjoinEleList;
        }
        public static double TotalAreaOfSolid (this Solid solid)
        {
            double totalArea = 0.0;
            FaceArray faces = solid.Faces;
            foreach (Face face in faces)
            {
                double faceArea = face.Area;
                totalArea += faceArea;
            }
            return totalArea;
        }
        public static bool CheckIfTwoSolidsAdjecent(this Solid solid1, Solid solid2)
        {
            if (solid1 == null || solid1.Volume == 0 || solid2 == null || solid2.Volume == 0)
            {
                return false;
            }
            Solid intersect = BooleanOperationsUtils.ExecuteBooleanOperation(solid1, solid2, BooleanOperationsType.Intersect);
            if (Math.Round(solid1.Volume, 3) == Math.Round(solid2.Volume, 3)  
                && Math.Round(solid1.Volume, 3) == Math.Round(intersect.Volume, 3))
            {
                return false;
            }
            bool isAdjecent = false;
            FaceArray solid1Faces = solid1.Faces;
            FaceArray solid2Faces = solid2.Faces;
            int solid1FacesNo = solid1Faces.Size;
            int solid2FacesNo = solid2Faces.Size;
            double solid1FacesArea = solid1.TotalAreaOfSolid();
            double solid2FacesArea = solid2.TotalAreaOfSolid();
            Solid solidUnion = null;
            if ((solid1 != null) && (solid1.Volume > 0) 
                && (solid2 != null) && (solid2.Volume > 0))
            {
                solidUnion = BooleanOperationsUtils.ExecuteBooleanOperation(solid1, solid2, BooleanOperationsType.Union);
            }
            if ((solidUnion != null) && (solidUnion.Volume > 0))
            {
                int solidUnionFacesNo = solidUnion.Faces.Size;
                double solidUnionFacesArea = solidUnion.TotalAreaOfSolid();
                if ((solidUnionFacesNo < (solid1FacesNo + solid2FacesNo)) ||
                    (Math.Round(solidUnionFacesArea, 2) != Math.Round(solid1FacesArea + solid2FacesArea, 2)))
                {
                    isAdjecent = true;
                }
            }
            return isAdjecent;
        }
        public static void JoinElements(this List<Element> eles, Document doc)
        {
            for (int i = 0; i < eles.Count; i++)
            {
                Solid ele1Solid = eles[i].UnionAllElementSolids();
                for (int j = i+1; j < eles.Count; j++)
                {
                    Solid ele2Solid = eles[j].UnionAllElementSolids();
                    bool isJoined = JoinGeometryUtils.AreElementsJoined(doc, eles[i], eles[j]);
                    if ((eles[i].Id != eles[j].Id) && (ele1Solid.CheckIfTwoSolidsAdjecent(ele2Solid)) && (isJoined is false))
                    {
                        JoinGeometryUtils.JoinGeometry(doc, eles[i], eles[j]);
                    }
                }
            }
        }
        public static (double, double, List<Wall>) GetPitWall (this Floor pitBase, Document doc)
        {
            Debug.WriteLine("GetPitWall: " + pitBase.Name);
            double pitWallThickness = 0.0;
            double pitWallMaxHeight = 0.0;
            List<Wall> pitWallList = new List<Wall>();
            Solid pitBaseSolid = pitBase.GetAllSolidsAdvance(true).First();
            var pitBaseName = pitBase.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();
            string pitName = pitBaseName;
            if (pitBaseName.Contains("の"))
            {
                pitName = pitBaseName.Substring(0, pitBaseName.IndexOf("の"));
            }
            var pitWallColl = new FilteredElementCollector(doc).WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Walls).OfClass(typeof(Wall)).Cast<Wall>()
                .Where(e => (e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() != null) 
                    && (e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains(pitName)));
            if (pitWallColl.Count() > 0)
            {
                int i = 0;
                foreach (Wall pitWall in pitWallColl)
                {
                    Debug.WriteLine(i + " GetPitWall: " + pitBase.Name + " | " + pitWall.Id);
                    i++;
                    Solid pitWallSolid = pitWall.GetAllSolidsAdvance(true).First();
                    bool isAdjacent = pitBaseSolid.CheckIfTwoSolidsAdjecent(pitWallSolid);
                    if (isAdjacent is true)
                    {
                        pitWallThickness = pitWall.WallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble();
                        var pitWallHeight = pitWall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
                        if (pitWallMaxHeight < pitWallHeight)
                        {
                            pitWallMaxHeight = pitWallHeight;
                        }
                        pitWallList.Add(pitWall);
                    }
                }
            }
            if (pitWallThickness == 0.0)
            {
                pitWallThickness = 150 / 304.8;
            }
            return (pitWallThickness, pitWallMaxHeight, pitWallList);
        }
        public static FamilyInstance GetJibanOnPit (this Floor pitBase, Document doc)
        {
            FamilyInstance jibanInst = null;
            if (pitBase is null)
            {
                return null;
            }
            Solid pitBaseSolid = pitBase.GetAllSolidsAdvance(true).First();
            Face pitBaseBotFace = pitBaseSolid.GetBottomPlanarFace();
            Solid pitBaseBotExtrudeUp = pitBaseBotFace.ExtrudeFace(XYZ.BasisZ, 50);

            var jibanColl = new FilteredElementCollector(doc).WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_CurtainWallPanels)
                .OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList();

            if (jibanColl.Count > 0)
            {
                foreach (var jiban in jibanColl)
                {
                    var jibanExtrudeDn = jiban.JibanExtrude(XYZ.BasisZ * -1, 50);
                    if (jibanExtrudeDn != null)
                    {
                        var xSolid = BooleanOperationsUtils.ExecuteBooleanOperation(pitBaseBotExtrudeUp, jibanExtrudeDn, BooleanOperationsType.Intersect);
                        if ((xSolid != null) && (xSolid.Volume > 0))
                        {
                            jibanInst = jiban;
                            break;
                        }
                    }
                }
            }
            return jibanInst;
        }
        public static Solid ExtrudeFace(this Face face, XYZ dir, double dist)
        {
            if (face is null)
            {
                return null;
            }
            var cls = face.GetEdgesAsCurveLoops();
            var maxCls = cls.GetMaxCurveLoop();
            List<CurveLoop> clList = new List<CurveLoop>();
            clList.Add(maxCls);
            Solid extrude = GeometryCreationUtilities.CreateExtrusionGeometry(clList, dir, dist);
            return extrude;
        }
        public static Solid ExtrudeCurveLoop(this IList<CurveLoop> curveLoopList, XYZ dir, double dist)
        {
            if ((curveLoopList is null)
                || (curveLoopList.Count == 0))
            {
                return null;
            }
            Solid extrude = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoopList, dir, dist);
            return extrude;
        }
        public static Solid ExtrudeFaceUpAndDn (this Face face, double dist)
        {
            Solid solidUp = face.ExtrudeFace(XYZ.BasisZ, dist);
            Solid solidDn = face.ExtrudeFace(XYZ.BasisZ * -1, dist);
            Solid solidUpAndDn = BooleanOperationsUtils.ExecuteBooleanOperation(solidUp, solidDn, BooleanOperationsType.Union);
            return solidUpAndDn;
        }
        public static Solid ExtrudeCurveLoopUpAndDn(this IList<CurveLoop> curveLoopList, double dist)
        {
            Solid solidUp = curveLoopList.ExtrudeCurveLoop(XYZ.BasisZ, dist);
            Solid solidDn = curveLoopList.ExtrudeCurveLoop(XYZ.BasisZ * -1, dist);
            Solid solidUpAndDn = BooleanOperationsUtils.ExecuteBooleanOperation(solidUp, solidDn, BooleanOperationsType.Union);
            return solidUpAndDn;
        }
        public static (bool, XYZ, XYZ, XYZ, XYZ, double, double) CheckIfFaceRectangle(this Face flatFace)
        {
            bool isRectangle = false;
            XYZ rp1 = new XYZ();
            XYZ rp2 = new XYZ();
            XYZ rp3 = new XYZ();
            XYZ rp4 = new XYZ();
            double rdist1 = 0.0;
            double rdist2 = 0.0;

            var flatFaceCls = flatFace.GetEdgesAsCurveLoops();
            var flatFaceMaxCl = flatFaceCls.GetMaxCurveLoop();
            List<XYZ> noColinearCurvePoints = new List<XYZ>();
            if ((flatFaceMaxCl != null) && (flatFaceMaxCl.Count() > 4))
            {
                List<XYZ> curvePoints = new List<XYZ>();
                foreach (var flatFaceMaxCurve in flatFaceMaxCl)
                {
                    var point = flatFaceMaxCurve.GetEndPoint(0);
                    curvePoints.Add(point);
                    noColinearCurvePoints.Add(point);
                }
                for (int i = 0; i < curvePoints.Count; i++)
                {
                    var p1 = new XYZ();
                    var p2 = new XYZ();
                    var p3 = new XYZ();
                    if (i < curvePoints.Count - 2)
                    {
                        p1 = curvePoints[i];
                        p2 = curvePoints[i + 1];
                        p3 = curvePoints[i + 2];
                    }
                    else if (i == curvePoints.Count - 2)
                    {
                        p1 = curvePoints[i];
                        p2 = curvePoints[i + 1];
                        p3 = curvePoints[0];
                    }
                    else if (i == curvePoints.Count - 1)
                    {
                        p1 = curvePoints[i];
                        p2 = curvePoints[0];
                        p3 = curvePoints[1];
                    }
                    else break;
                    if ((p1 != null) && (p2 != null) && (p3 != null))
                    {
                        if (Utils.PointUtils.CheckColinearPoints(p1, p2, p3))
                        {
                            noColinearCurvePoints.Remove(p2);
                        }
                    }
                }
            }
            else if ((flatFaceMaxCl != null) && (flatFaceMaxCl.Count() == 4))
            {
                var p1 = flatFaceMaxCl.ElementAt(0).GetEndPoint(0);
                var p2 = flatFaceMaxCl.ElementAt(1).GetEndPoint(0);
                var p3 = flatFaceMaxCl.ElementAt(2).GetEndPoint(0);
                var p4 = flatFaceMaxCl.ElementAt(3).GetEndPoint(0);
                noColinearCurvePoints.Add(p1);
                noColinearCurvePoints.Add(p2);
                noColinearCurvePoints.Add(p3);
                noColinearCurvePoints.Add(p4);
            }
            if ((noColinearCurvePoints != null) && (noColinearCurvePoints.Count() == 4))
            {
                var p1 = noColinearCurvePoints[0];
                var p2 = noColinearCurvePoints[1];
                var p3 = noColinearCurvePoints[2];
                var p4 = noColinearCurvePoints[3];

                var dist1 = p1.DistanceTo(p2);
                var dist2 = p2.DistanceTo(p3);
                var dist3 = p3.DistanceTo(p4);
                var dist4 = p4.DistanceTo(p1);

                Line line1 = Line.CreateBound(p1, p2);
                Line line2 = Line.CreateBound(p2, p3);

                if ((Math.Round(dist1, 3) == Math.Round(dist3, 3)) && (Math.Round(dist2, 3) == Math.Round(dist4, 3)))
                {
                    isRectangle = true;
                    if (dist1 >= dist2)
                    {
                        rp1 = p2;
                        rp2 = p3;
                        rp3 = p4;
                        rp4 = p1;
                        rdist1 = dist2;
                        rdist2 = dist1;
                    }
                    else
                    {
                        rp1 = p1;
                        rp2 = p2;
                        rp3 = p3;
                        rp4 = p4;
                        rdist1 = dist1;
                        rdist2 = dist2;
                    }
                }
            }
            return (isRectangle, rp1, rp2, rp3, rp4, rdist1, rdist2);
        }
        public static (bool, XYZ, XYZ, XYZ, XYZ, double, double) CheckIfSolidRectangle (this Solid solid)
        {
            bool isRectangle = false;
            XYZ rp1 = new XYZ();
            XYZ rp2 = new XYZ();
            XYZ rp3 = new XYZ();
            XYZ rp4 = new XYZ();
            double rdist1 = 0.0;
            double rdist2 = 0.0;

            var flatFace = solid.GetFlatFaceOfSolid();
            var flatFaceCls = flatFace.GetEdgesAsCurveLoops();
            var flatFaceMaxCl = flatFaceCls.GetMaxCurveLoop();
            List<XYZ> noColinearCurvePoints = new List<XYZ>();
            if ((flatFaceMaxCl != null) && (flatFaceMaxCl.Count() > 4))
            {
                List<XYZ> curvePoints = new List<XYZ>();
                foreach (var flatFaceMaxCurve in flatFaceMaxCl)
                {
                    var point = flatFaceMaxCurve.GetEndPoint(0);
                    curvePoints.Add(point);
                    noColinearCurvePoints.Add(point);
                }
                for (int i = 0; i < curvePoints.Count; i++)
                {
                    var p1 = new XYZ();
                    var p2 = new XYZ();
                    var p3 = new XYZ();
                    if (i < curvePoints.Count - 2)
                    {
                        p1 = curvePoints[i];
                        p2 = curvePoints[i + 1];
                        p3 = curvePoints[i + 2];
                    }
                    else if (i == curvePoints.Count - 2)
                    {
                        p1 = curvePoints[i];
                        p2 = curvePoints[i + 1];
                        p3 = curvePoints[0];
                    }
                    else if (i == curvePoints.Count - 1)
                    {
                        p1 = curvePoints[i];
                        p2 = curvePoints[0];
                        p3 = curvePoints[1];
                    }
                    else break;
                    if ((p1 != null) && (p2 != null) && (p3 != null))
                    {
                        if (Utils.PointUtils.CheckColinearPoints(p1, p2, p3))
                        {
                            noColinearCurvePoints.Remove(p2);
                        }
                    }
                }
            }
            else if ((flatFaceMaxCl != null) && (flatFaceMaxCl.Count() == 4))
            {
                var p1 = flatFaceMaxCl.ElementAt(0).GetEndPoint(0);
                var p2 = flatFaceMaxCl.ElementAt(1).GetEndPoint(0);
                var p3 = flatFaceMaxCl.ElementAt(2).GetEndPoint(0);
                var p4 = flatFaceMaxCl.ElementAt(3).GetEndPoint(0);
                noColinearCurvePoints.Add(p1);
                noColinearCurvePoints.Add(p2);
                noColinearCurvePoints.Add(p3);
                noColinearCurvePoints.Add(p4);
            }
            if ((noColinearCurvePoints != null) && (noColinearCurvePoints.Count() == 4))
            {
                var p1 = noColinearCurvePoints[0];
                var p2 = noColinearCurvePoints[1];
                var p3 = noColinearCurvePoints[2];
                var p4 = noColinearCurvePoints[3];

                var dist1 = p1.DistanceTo(p2);
                var dist2 = p2.DistanceTo(p3);
                var dist3 = p3.DistanceTo(p4);
                var dist4 = p4.DistanceTo(p1);

                Line line1 = Line.CreateBound(p1, p2);
                Line line2 = Line.CreateBound(p2, p3);

                if ((Math.Round(dist1, 3) == Math.Round(dist3, 3)) && (Math.Round(dist2, 3) == Math.Round(dist4, 3)))
                {
                    isRectangle = true;
                    if (dist1 >= dist2)
                    {
                        rp1 = p2;
                        rp2 = p3;
                        rp3 = p4;
                        rp4 = p1;
                        rdist1 = dist2;
                        rdist2 = dist1;
                    }
                    else
                    {
                        rp1 = p1;
                        rp2 = p2;
                        rp3 = p3;
                        rp4 = p4;
                        rdist1 = dist1;
                        rdist2 = dist2;
                    }
                }
            }
            return (isRectangle, rp1, rp2, rp3, rp4, rdist1, rdist2);
        }
        public static (bool, XYZ, XYZ, XYZ, XYZ, double, double) CheckIfFloorRectangle(this Floor floor, Document doc)
        {
            XYZ rp1 = new XYZ();
            XYZ rp2 = new XYZ();
            XYZ rp3 = new XYZ();
            XYZ rp4 = new XYZ();
            double rdist1 = 0.0;
            double rdist2 = 0.0;

            List<Element> otherFloorColl = new FilteredElementCollector(doc).WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Floors).OfClass(typeof(Floor)).Cast<Element>().ToList();

            var floorSolid = floor.GetAllSolidsAdvance().First();
            var floorSolidRecCheck = floorSolid.CheckIfSolidRectangle();

            bool isRectangle = floorSolidRecCheck.Item1;
            if (isRectangle is true)
            {
                XYZ p1 = floorSolidRecCheck.Item2;
                XYZ p2 = floorSolidRecCheck.Item3;
                XYZ p3 = floorSolidRecCheck.Item4;
                XYZ p4 = floorSolidRecCheck.Item5;
                double dist1 = floorSolidRecCheck.Item6;
                double dist2 = floorSolidRecCheck.Item7;

                var floorSpanAngle = floor.SpanDirectionAngle;
                XYZ horVec = new XYZ(1, 0, 0);
                Transform tfSpanAngle = Transform.CreateRotation(XYZ.BasisZ, floorSpanAngle);
                XYZ spanVec = tfSpanAngle.OfVector(horVec);
                var shortLine = Line.CreateBound(p1, p2);
                if (Math.Round(shortLine.Direction.DotProduct(spanVec), 2) == 0.00)
                {
                    rp1 = p2;
                    rp2 = p3;
                    rp3 = p4;
                    rp4 = p1;
                    rdist1 = dist2;
                    rdist2 = dist1;
                }
                else
                {
                    rp1 = p1;
                    rp2 = p2;
                    rp3 = p3;
                    rp4 = p4;
                    rdist1 = dist1;
                    rdist2 = dist2;
                }
            }

            return (isRectangle, rp1, rp2, rp3, rp4, rdist1, rdist2);
        }
        public static Solid CreateCylinderUpAndDn (this XYZ center, Document doc, double radius, double dist)
        {
            Level glLevel = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Levels).OfClass(typeof(Level))
                .Cast<Level>().First(e => e.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble().Equals(0));
            var glViewPlanId = glLevel.FindAssociatedPlanViewId();
            ViewPlan glViewPlan = doc.GetElement(glViewPlanId) as ViewPlan;
            var glSketchPlane = glViewPlan.SketchPlane;
            var glPlane = glSketchPlane.GetPlane();
            var halfCircle1 = Arc.Create(center, radius, 0, 180 * Math.PI / 180, XYZ.BasisX, XYZ.BasisY) as Curve;
            var halfCircle2 = Arc.Create(center, radius, 180 * Math.PI / 180, 360 * Math.PI / 180, XYZ.BasisX, XYZ.BasisY) as Curve;
            CurveLoop curveLoop = new CurveLoop();
            curveLoop.Append(halfCircle1);
            curveLoop.Append(halfCircle2);
            List<CurveLoop> curveLoopList = new List<CurveLoop>();
            curveLoopList.Add(curveLoop);

            Solid extrudeUp = GeometryCreationUtilities.CreateExtrusionGeometry(
                curveLoopList, XYZ.BasisZ, dist);
            Solid extrudeDn = GeometryCreationUtilities.CreateExtrusionGeometry(
                curveLoopList, XYZ.BasisZ * -1, dist);
            Solid extrudeUpAndDn = BooleanOperationsUtils.ExecuteBooleanOperation(
                extrudeUp, extrudeDn, BooleanOperationsType.Union);
            return extrudeUpAndDn;
        }

        public static Solid CreateCylinderUpAndDnByLevel(this XYZ center, Document doc, double radius, double dist, Level level)
        {
            var halfCircle1 = Arc.Create(center, radius, 0, Math.PI, XYZ.BasisX, XYZ.BasisY) as Curve;
            var halfCircle2 = Arc.Create(center, radius, Math.PI, 2 * Math.PI, XYZ.BasisX, XYZ.BasisY) as Curve;

            CurveLoop curveLoop = new CurveLoop();
            curveLoop.Append(halfCircle1);
            curveLoop.Append(halfCircle2);

            List<CurveLoop> curveLoopList = new List<CurveLoop> { curveLoop };

            Solid extrudeUp = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoopList, XYZ.BasisZ, dist);
            Solid extrudeDn = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoopList, -XYZ.BasisZ, dist);

            Solid extrudeUpAndDn = BooleanOperationsUtils.ExecuteBooleanOperation(extrudeUp, extrudeDn, BooleanOperationsType.Union);

            double elevation = level.Elevation;
            BoundingBoxXYZ boundingBox = extrudeUpAndDn.GetBoundingBox();
            double bottomZ = boundingBox.Min.Z;
            double translationZ = elevation - bottomZ;

            Transform translationTransform = Transform.CreateTranslation(new XYZ(0, 0, translationZ));
            Solid translatedSolid = SolidUtils.CreateTransformed(extrudeUpAndDn, translationTransform);

            

            return translatedSolid;
        }
        public static bool CheckIfListContainsElement (this List<Element> listEle, Element checkEle)
        {
            bool isIncluded = false;
            if ((listEle is null) && (listEle.Count == 0))
            {
                return false;
            }
            foreach (var ele in listEle)
            {
                if (ele.Id == checkEle.Id)
                {
                    isIncluded = true;
                    break;
                }
            }
            return isIncluded;
        }
        public static (bool, List<Face>, double, double) CheckIfPrism (this Solid solid)
        {
            if ((solid is null) || (solid.Volume == 0))
            {
                return (false, null, 0, 0);
            }
            List<Face> UpDnFaces = new List<Face>();
            FaceArray faces = solid.Faces;
            int countUpFace = 0;
            int countDnFace = 0;
            Face upFace = null;
            Face dnFace = null;
            List<Face> upFaces = new List<Face>();
            List<Face> dnFaces = new List<Face>();
            for (int i = 0; i < faces.Size; i++)
            {
                var curFace = faces.get_Item(i) as PlanarFace;
                if (curFace != null)
                {
                    var normal = curFace.FaceNormal.Z;
                    if (normal == 1)
                    {
                        countUpFace++;
                        upFace = curFace;
                        upFaces.Add(curFace);
                    }
                    else if (normal == -1)
                    {
                        countDnFace++;
                        dnFace = curFace;
                        dnFaces.Add(curFace);
                    }
                }
            }
            if (countUpFace > 1)
            {
                var upFaceLevel = upFaces.GroupBy(f => Math.Round((f as PlanarFace).Origin.Z, 3)).Count();
                if (upFaceLevel == 1)
                {
                    countUpFace = 1;
                }
            }
            if (countDnFace > 1)
            {
                var dnFaceLevel = dnFaces.GroupBy(f => Math.Round((f as PlanarFace).Origin.Z, 3)).Count();
                if (dnFaceLevel == 1)
                {
                    countDnFace = 1;
                }
            }
            if ((countUpFace == 1) && (countDnFace == 1))
            {
                UpDnFaces.Add(upFace);
                UpDnFaces.Add(dnFace);
                Plane upPlanarFace = Plane.CreateByNormalAndOrigin((upFace as PlanarFace).FaceNormal, (upFace as PlanarFace).Origin);
                Plane dnPlanarFace = Plane.CreateByNormalAndOrigin((dnFace as PlanarFace).FaceNormal, (dnFace as PlanarFace).Origin);
                UV uv = new UV();
                double dist = 0.0;
                dnPlanarFace.Project(upPlanarFace.Origin, out uv, out dist);
                double area = solid.Volume / dist;
                return (true, UpDnFaces, dist, area);
            }
            else return (false, null, 0, 0);
        }
        public static bool ElementOverlap(this Solid solid, List<FamilyInstance> eleList)
        {
            if ((solid == null) || (eleList == null))
            {
                return false;
            }
            foreach (var ele in eleList)
            {
                var eleSolidList = ele.GetAllSolidsAdvance(true).ToList();
                if (eleSolidList.Count > 0)
                {
                    foreach (var eleSolid in eleSolidList)
                    {
                        var xSolid = BooleanOperationsUtils.ExecuteBooleanOperation(solid, eleSolid, BooleanOperationsType.Intersect);
                        if ((xSolid != null) && (xSolid.Volume > 0))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        public static double SolidVerFaceArea(this Solid solid)
        {
            if ((solid == null) || (solid.Volume <= 0))
            {
                return 0;
            }
            double verFacesArea = 0.0;
            var faces = solid.Faces;
            foreach (Face face in faces)
            {
                if ((face is PlanarFace && Math.Round((face as PlanarFace).FaceNormal.Z, 3) == 0) 
                    || (face is CylindricalFace))
                {
                    verFacesArea += face.Area;
                }
            }
            return verFacesArea;
        }
        public static Solid UnionSolidList(this List<Solid> solidList)
        {
            if ((solidList == null) || (solidList.Count == 0))
            {
                return null;
            }
            Solid unionSolid = null;
            foreach (var solid in solidList)
            {
                if (unionSolid == null)
                {
                    unionSolid = solid;
                }
                else
                {
                    unionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(unionSolid, solid, BooleanOperationsType.Union);
                }
            }
            return unionSolid;
        }
        public static Solid ExtrudeFromTwoLines(this Curve line1, Curve line2, XYZ dir, double dist)
        {
            if ((line1 is null) || (line2 is null))
            {
                return null;
            }
            var p1 = line1.GetEndPoint(0);
            var p2 = line1.GetEndPoint(1);
            var p3 = line2.GetEndPoint(1);
            var p4 = line2.GetEndPoint(0);

            Line l1 = Line.CreateBound(p1, p2);
            Line l2 = Line.CreateBound(p2, p3);
            Line l3 = Line.CreateBound(p3, p4);
            Line l4 = Line.CreateBound(p4, p1);

            List<Curve> curveList = new List<Curve>() { l1, l2, l3, l4};
            CurveLoop cl = CurveLoop.Create(curveList);
            List<CurveLoop> cls = new List<CurveLoop>() { cl };
            return GeometryCreationUtilities.CreateExtrusionGeometry(cls, dir, dist);
        }
        public static void CreateFormworkBySolid(this Solid solid, Application app, Document doc, ElementId fkTypeId, double fkThickness, ElementId glId,
            bool overrideHeight = false, double inputFkHeight = 0, double inputBaseOffset = 0)
        {
            if ((solid != null) && (solid.Volume > 0))
            {
                List<Solid> afterCutSolidList = solid.SplitSolidByItsOwnVerticalFaces(doc);
                if ((afterCutSolidList != null) && (afterCutSolidList.Count > 0))
                {
                    foreach (var afterCutSolid in afterCutSolidList)
                    {
                        var botFace = afterCutSolid.GetBottomPlanarFaceList()
                            .OrderBy(f => f.Origin.Z).First();
                        var topFace = afterCutSolid.GetTopPlanarFaceList()
                            .OrderByDescending(f => f.Origin.Z).First();
                        if ((botFace != null) && (topFace != null))
                        {
                            double fkHeight = topFace.Origin.Z - botFace.Origin.Z;
                            if (overrideHeight)
                            {
                                fkHeight = inputFkHeight;
                            }
                            double baseOffset = botFace.Origin.Z;
                            if (overrideHeight)
                            {
                                baseOffset = inputBaseOffset;
                            }
                            //topFace.ExtrudeFace(XYZ.BasisZ, 10).BakeSolidToDirectShape(doc);
                            //botFace.ExtrudeFace(XYZ.BasisZ * -1, 1).BakeSolidToDirectShape(doc);
                            var curveLoop = botFace.GetEdgesAsCurveLoops().First();
                            List<CurveLoop> cls = new List<CurveLoop>() { curveLoop };
                            //cls.ExtrudeCurveLoop(XYZ.BasisZ, 10).BakeSolidToDirectShape(doc);
                            var wallLocLine = curveLoop.GetCenterLineOfCurveLoop(app, doc, fkThickness);
                            if ((wallLocLine != null) && (wallLocLine.Length > app.ShortCurveTolerance))
                            {
                                try
                                {
                                    if (fkHeight >= 50 / 304.8)
                                    {
                                        var fkWall = Wall.Create(doc, wallLocLine, fkTypeId, glId, fkHeight, 0, false, false);
                                        fkWall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(baseOffset);
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }
            }
        }
        public static List<Solid> SplitSolidByItsOwnVerticalFaces(this Solid solid, Document doc)
        {
            if ((solid == null) || (solid.Volume == 0))
            {
                return null;
            }
            List<Solid> prismSolidList = new List<Solid>();
            Solid[] beCutSolidList = new Solid[] { solid };
            while (beCutSolidList.Count() > 0)
            {
                var beCutSolid = beCutSolidList.ElementAt(0);
                List<Face> verFaceList = beCutSolid.GetPerVerFaces(20 / 304.8);
                bool overCount = false;
                if (verFaceList.Count > 2)
                {
                    var orderedFaceList = verFaceList.OrderBy(f => (f as PlanarFace).Origin.X)
                                                .ThenBy(f => (f as PlanarFace).Origin.Y);
                    int i = 0;
                    while (i < verFaceList.Count)
                    {
                        var cutFace = orderedFaceList.ElementAt(i);
                        //cutFace.ExtrudeFace((cutFace as PlanarFace).FaceNormal, 10).BakeSolidToDirectShape(doc);
                        var cutPlane = Plane.CreateByNormalAndOrigin((cutFace as PlanarFace).FaceNormal, (cutFace as PlanarFace).Origin);
                        var revCutPlane = Plane.CreateByNormalAndOrigin((cutFace as PlanarFace).FaceNormal * -1, (cutFace as PlanarFace).Origin);
                        var cutSolid = BooleanOperationsUtils.CutWithHalfSpace(beCutSolid, cutPlane);
                        var revCutSolid = BooleanOperationsUtils.CutWithHalfSpace(beCutSolid, revCutPlane);
                        if ((cutSolid != null) && (cutSolid.Volume > 0)
                            && (revCutSolid != null) && (revCutSolid.Volume > 0))
                        {
                            //cutSolid.BakeSolidToDirectShape(doc);
                            //revCutSolid.BakeSolidToDirectShape(doc);
                            var cutSolidVerFaceList = cutSolid.GetPerVerFaces(20 / 304.8);
                            var revCutSolidVerFaceList = revCutSolid.GetPerVerFaces(20 / 304.8);
                            beCutSolidList = beCutSolidList.Where((source, index) => index != 0).ToArray();
                            if ((cutSolidVerFaceList != null) && (cutSolidVerFaceList.Count > 2))
                            {
                                beCutSolidList = new Solid[beCutSolidList.Count() + 1];
                                beCutSolidList[beCutSolidList.Count() - 1] = cutSolid;
                            }
                            else
                            {
                                prismSolidList.Add(cutSolid);
                            }
                            if ((revCutSolidVerFaceList != null) && (revCutSolidVerFaceList.Count > 2))
                            {
                                beCutSolidList = new Solid[beCutSolidList.Count() + 1];
                                beCutSolidList[beCutSolidList.Count() - 1] = revCutSolid;
                            }
                            else
                            {
                                prismSolidList.Add(revCutSolid);
                            }
                            i = verFaceList.Count + 1;
                        }
                        else
                        {
                            i++;
                        }
                    }
                    if (i == verFaceList.Count())
                    {
                        overCount = true;
                    }
                }
                if ((verFaceList.Count <= 2) || (overCount == true))
                {
                    beCutSolidList = beCutSolidList.Where((source, index) => index != 0).ToArray();
                    prismSolidList.Add(beCutSolid);
                }
                beCutSolidList = beCutSolidList.RemoveNullSolidInArray();
            }
            return prismSolidList;
        }
        public static Solid[] RemoveNullSolidInArray(this Solid[] solidArr)
        {
            Solid[] newSolidArr = new Solid[0];
            foreach (var solid in solidArr)
            {
                if ((solid != null) && (solid.Volume > 0))
                {
                    newSolidArr = new Solid[newSolidArr.Count() + 1];
                    newSolidArr[newSolidArr.Count() - 1] = solid;
                }
            }
            return newSolidArr;
        }
        public static Solid GetFoundBotFaceExtrudeToGl(this List<FamilyInstance> foundColl)
        {
            if (foundColl == null || foundColl.Count == 0) return null;
            Solid union = null;
            foreach (var found in foundColl)
            {
                var foundationThickness = found.Symbol.get_Parameter(BuiltInParameter.STRUCTURAL_FOUNDATION_THICKNESS).AsDouble();
                var foundTopOffset = found.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble();
                var foundSolid = found.FoundMainAndYoboriSolid();
                var foundBotFace = foundSolid.GetBottomPlanarFace();
                var foundCls = foundBotFace.GetEdgesAsCurveLoops();
                var foundBotExtrude = GeometryCreationUtilities.CreateExtrusionGeometry(foundCls, XYZ.BasisZ, Math.Abs(foundTopOffset) + foundationThickness);
                union = (union == null || union.Volume == 0) ? foundBotExtrude 
                    : BooleanOperationsUtils.ExecuteBooleanOperation(union, foundBotExtrude, BooleanOperationsType.Union);
            }
            return union;
        }
        public static double GetSolidHeight(this Solid solid)
        {
            if (solid == null || solid.Volume == 0)
            {
                return 0.0;
            }

            // Get the top and bottom planar faces
            PlanarFace topFace = solid.GetTopPlanarFace();
            PlanarFace bottomFace = solid.GetBottomPlanarFace();

            // Ensure both top and bottom faces exist
            if (topFace == null || bottomFace == null)
            {
                return 0.0;
            }

            // Calculate the height as the difference between top and bottom face Z values
            double topZ = topFace.Origin.Z;
            double bottomZ = bottomFace.Origin.Z;

            double height = topZ - bottomZ;

            return Math.Abs(height); // Return the absolute value to avoid negative height
        }

    }
}
