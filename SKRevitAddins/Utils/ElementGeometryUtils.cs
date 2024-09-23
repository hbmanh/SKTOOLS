#region Namespaces

using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

#endregion

namespace SKRevitAddins.Utils
{
    public static class ElementGeometryUtils
    {
        /// <summary>
        /// Get Solid of FamilyInstance by Symbol
        /// </summary>
        /// <param name="ele"></param>
        /// <param name="transformSolid">Return solid in model coordinate when trnasformSolid==true</param>
        /// <returns></returns>
        public static List<Solid> GetSolidsBySymbol(this Element ele, bool transformSolid = false)
        {
            var solids = new List<Solid>();
            var op = new Options();
            op.ComputeReferences = true;
            op.IncludeNonVisibleObjects = true;
            op.DetailLevel = ViewDetailLevel.Undefined;
            var geoE = ele.get_Geometry(op);
            if (geoE == null) return solids;
            foreach (var geoO in geoE)
            {
                var geoI = geoO as GeometryInstance;
                if (geoI == null) continue;
                var instanceGeoE = geoI.GetSymbolGeometry();
                var tf = geoI.Transform;
                foreach (var instanceGeoObj in instanceGeoE)
                {
                    var solid1 = instanceGeoObj as Solid;
                    Solid solid = solid1;
                    if (transformSolid)
                    {
                        solid = SolidUtils.CreateTransformed(solid1, tf);
                        solids.Add(solid);
                    }
                    if (solid == null || solid.Faces.Size == 0) continue;

                }
            }
            return solids;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ele"></param>
        /// <returns></returns>
        public static List<Solid> GetSolids(this Element ele)
        {
            List<Solid> solids = new List<Solid>();
            // Option :
            var op = new Options();
            op.ComputeReferences = true;
            op.IncludeNonVisibleObjects = true;
            op.DetailLevel = ViewDetailLevel.Undefined;
            var geoE = ele.get_Geometry(op);
            if (geoE == null) return solids;
            foreach (var geoO in geoE)
            {
                var solid = geoO as Solid;
                if (solid == null || solid.Faces.Size == 0 || solid.Edges.Size == 0)
                {
                    continue;
                }
                solids.Add(solid);
            }
            return solids;
        }
        public static List<Solid> GetAllSolids(this Element ele, bool transformSolid = false)
        {
            var solids = GetSolids(ele);
            if (solids.Count == 0)
            {
                solids = GetSolidsBySymbol(ele, transformSolid);
            }
            return solids;
        }

        public static List<Solid> GetAllSolidsAdvance(this Element instance, bool transformedSolid = false, View view = null)
        {
            List<Solid> solidList = new List<Solid>();
            if (instance == null)
                return solidList;
            GeometryElement geometryElement;
            if (instance.Category.Id.IntegerValue.Equals(-2003200))
            {
                geometryElement = instance.get_Geometry(new Options()
                {
                    ComputeReferences = true,
                    View = view
                });
            }
            else
            {
                geometryElement = instance.get_Geometry(new Options()
                {
                    ComputeReferences = true
                });
            }
            

            foreach (GeometryObject geometryObject1 in geometryElement)
            {
                GeometryInstance geometryInstance = geometryObject1 as GeometryInstance;
                if (null != geometryInstance)
                {
                    var tf = geometryInstance.Transform;
                    foreach (GeometryObject geometryObject2 in geometryInstance.GetSymbolGeometry())
                    {
                        Solid solid = geometryObject2 as Solid;
                        if (!(null == solid) && solid.Volume>0&& solid.Faces.Size != 0 && solid.Edges.Size != 0)
                        {
                            if (transformedSolid)
                            {
                                //solidList.Add(SolidUtils.CreateTransformed(solid, tf));
                                solid = SolidUtils.CreateTransformed(solid, tf);
                            }
                            solidList.Add(solid);
                        }
                    }
                }
                Solid solid1 = geometryObject1 as Solid;
                if (!(null == solid1) && solid1.Faces.Size != 0)
                    solidList.Add(solid1);
            }
            return solidList;
        }
        public static List<Solid> GetAllSolidsAdvanceFromMEPCurve(this MEPCurve mepCurve, bool transformedSolid = false, View view = null)
        {
            List<Solid> solidList = new List<Solid>();
            if (mepCurve == null)
                return solidList;

            LocationCurve locationCurve = mepCurve.Location as LocationCurve;
            if (locationCurve == null)
                return solidList;

            Curve curve = locationCurve.Curve;
            if (curve == null)
                return solidList;

            // Tạo ra một hình học Solid từ Curve của MEPCurve
            CurveLoop curveLoop = new CurveLoop();
            curveLoop.Append(curve);
            List<CurveLoop> curveLoops = new List<CurveLoop> { curveLoop };

            // Đặt chiều cao cho solid, có thể thay đổi cho phù hợp
            double height = UnitUtils.MmToFeet(100);

            try
            {
                Solid solid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoops, XYZ.BasisZ, height);
                solidList.Add(solid);
            }
            catch (Exception ex)
            {
                // Xử lý ngoại lệ nếu có
                TaskDialog.Show("Error", ex.Message);
            }

            return solidList;
        }



        public static List<Solid> GetOriginalSolids(this FamilyInstance instance, bool transformedSolid = false, View view = null)
        {
            List<Solid> solidList = new List<Solid>();
            if (instance == null)
                return solidList;
            GeometryElement geometryElement = instance.GetOriginalGeometry(new Options()
            {
                ComputeReferences = true
            });

            foreach (GeometryObject geometryObject1 in geometryElement)
            {
                GeometryInstance geometryInstance = geometryObject1 as GeometryInstance;
                if (null != geometryInstance)
                {
                    var tf = geometryInstance.Transform;
                    foreach (GeometryObject geometryObject2 in geometryInstance.GetSymbolGeometry())
                    {
                        Solid solid = geometryObject2 as Solid;
                        if (!(null == solid) && solid.Faces.Size != 0 && solid.Edges.Size != 0)
                        {
                            if (transformedSolid)
                            {
                                //solidList.Add(SolidUtils.CreateTransformed(solid, tf));
                                solid = SolidUtils.CreateTransformed(solid, tf);
                            }
                            solidList.Add(solid);
                        }
                    }
                }
                Solid solid1 = geometryObject1 as Solid;
                if (!(null == solid1) && solid1.Faces.Size != 0)
                    solidList.Add(solid1);
            }
            return solidList;
        }

        public static Solid UnionAllElementSolids (this Element ele)
        {
            Solid solidUnion = null;
            List<Solid> solids = ele.GetAllSolidsAdvance(true);
            foreach (var solid in solids)
            {
                if (solidUnion is null)
                {
                    solidUnion = solid;
                }
                else
                {
                    solidUnion = BooleanOperationsUtils.ExecuteBooleanOperation(solidUnion, solid, BooleanOperationsType.Union);
                }
            }
            return solidUnion;
        }

        public static List<Element> GetJoinedElementList(this Element ele, List<Element> checkingEleList, Document doc)
        {
            if ((checkingEleList is null) || (checkingEleList.Count <= 0))
            {
                return null;
            }
            List<Element> joinedEleList = new List<Element>();
            foreach (var checkingEle in checkingEleList)
            {
                bool isJoined = false;
                isJoined = JoinGeometryUtils.AreElementsJoined(doc, ele, checkingEle);
                if (isJoined is true)
                {
                    joinedEleList.Add(checkingEle);
                }
            }
            return joinedEleList;
        }

        public static List<Face> GetElementVerticalFaces(this Element ele)
        {
            List<Face> listVerFace = new List<Face>();
            Solid eleSolid = ele.UnionInstanceSolids();
            FaceArray eleFaces = eleSolid.Faces;
            for (int i = 0; i < eleFaces.Size; i++)
            {
                Face curFace = eleFaces.get_Item(i);
                double curFaceNormalZ = (curFace as PlanarFace).FaceNormal.Z;
                if (curFaceNormalZ.Equals(0))
                {
                    listVerFace.Add(curFace);
                }
            }
            return listVerFace;
        }

        public static List<Face> GetSolidVerticalFaces(this Solid solid)
        {
            if ((solid == null) || (solid.Volume == 0))
            {
                return null;
            }
            List<Face> listVerFace = new List<Face>();
            FaceArray eleFaces = solid.Faces;
            for (int i = 0; i < eleFaces.Size; i++)
            {
                Face curFace = eleFaces.get_Item(i);
                double curFaceNormalZ = (curFace as PlanarFace).FaceNormal.Z;
                if (Math.Round(curFaceNormalZ, 3) == 0)
                {
                    listVerFace.Add(curFace);
                }
            }
            return listVerFace;
        }
        public static List<Face> GetPerVerFaces(this Solid solid, double dist)
        {
            if ((solid == null) || (solid.Volume == 0))
            {
                return null;
            }
            List<Face> tempVerFaceList = solid.GetSolidVerticalFaces();
            List<Face> verFaceList = new List<Face>();
            int removeIndex1 = -1;
            int removeIndex2 = -2;
            if (tempVerFaceList.Count > 0)
            {
                for (int i = 0; i < tempVerFaceList.Count; i++)
                {
                    var iFace = tempVerFaceList[i] as PlanarFace;
                    var iPlane = Plane.CreateByNormalAndOrigin(iFace.FaceNormal, iFace.Origin);
                    for (int j = i + 1; j < tempVerFaceList.Count; j++)
                    {
                        var jFace = tempVerFaceList[j] as PlanarFace;
                        var jPlane = Plane.CreateByNormalAndOrigin(jFace.FaceNormal, jFace.Origin);
                        if (iFace.FaceNormal.AngleTo(jFace.FaceNormal) == Math.PI)
                        {
                            UV uv;
                            double curDist;
                            jPlane.Project(iFace.Origin, out uv, out curDist);
                            if (curDist <= dist)
                            {
                                removeIndex1 = i;
                                removeIndex2 = j;
                            }
                        }
                    }
                }
                for (int k = 0; k < tempVerFaceList.Count; k++)
                {
                    if ((k != removeIndex1) && (k != removeIndex2))
                    {
                        var face = tempVerFaceList[k];
                        verFaceList.Add(face);
                    }
                }
            }
            return verFaceList;
        }

        public static List<Floor> RemoveFloorInFloorList(this List<Floor> floorList, Floor removedFloor)
        {
            List<Floor> newFloorList = new List<Floor>();
            foreach (var floor in floorList)
            {
                if (floor.Id != removedFloor.Id)
                {
                    newFloorList.Add(floor);
                }
            }
            return newFloorList;
        }


        public class WallComparer : EqualityComparer<Wall>
        {
            public override bool Equals(Wall w1, Wall w2)
            {
                bool sameWall = w1.Id.IntegerValue == w2.Id.IntegerValue;
                return sameWall;
            }

            public override int GetHashCode(Wall obj)
            {
                return 0;
            }
        }
       
    }
}
