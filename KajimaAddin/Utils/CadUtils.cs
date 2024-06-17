#region Namespaces
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using Autodesk.Revit.DB;
#endregion

namespace SKToolsAddins
{
    public static class CadUtils
    {
        /// <summary>
        /// Lấy về tất cả Solid có Faces.Count > 0 của file CAD.
        /// </summary>
        /// <param name="cadInstance"></param>
        /// <param name="layerName"></param>
        /// <returns></returns>


        public static List<string> GetAllLayer(this ImportInstance cadInstance)
        {
            List<string> allLayers = new List<string>();

            // Defaults to medium detail, no references and no view.
            Options option = new Options();
            option.IncludeNonVisibleObjects = true;
            option.ComputeReferences = true;
            option.DetailLevel = ViewDetailLevel.Fine;

            // option.View = cadInstance.Document.ActiveView;
            GeometryElement geoElement = cadInstance.get_Geometry(option);

            foreach (GeometryObject geoObject in geoElement)
            {
                if (!(geoObject is GeometryInstance)) continue;
                GeometryInstance geoInstance = geoObject as GeometryInstance;
                GeometryElement geoElement2 = geoInstance.GetInstanceGeometry();

                foreach (GeometryObject geoObject2 in geoElement2)
                {
                    if (geoObject2 is Solid)
                    {
                        Solid solid = geoObject2 as Solid;
                        if (solid.Volume < 0.001) continue;

                        FaceArray faceArray = solid.Faces;
                        foreach (Face face in faceArray)
                        {
                            ElementId elementId = face.GraphicsStyleId;
                            GraphicsStyle graphicsStyle =
                                cadInstance.Document.GetElement(elementId)
                                    as GraphicsStyle;
                            Category styleCategory = graphicsStyle.GraphicsStyleCategory;
                            allLayers.Add(styleCategory.Name);
                        }
                    }
                    else
                    {
                        ElementId elementId = geoObject2.GraphicsStyleId;
                        GraphicsStyle graphicsStyle =
                            cadInstance.Document.GetElement(elementId)
                                as GraphicsStyle;
                        Category styleCategory = graphicsStyle.GraphicsStyleCategory;
                        allLayers.Add(styleCategory.Name);
                    }
                }
            }

            allLayers = allLayers.Distinct().ToList();
            allLayers.Sort();
            return allLayers;
        }
        public static string GetLayerNameFromCurveOrPolyline(GeometryObject geometryObject, ImportInstance cadInstance)
        {
            string layerName = "";

            // Kiểm tra xem geometryObject có phải là đường cong hay không
            if (geometryObject is Curve curve)
            {
                // Get the graphics style of the curve
                GraphicsStyle graphicsStyle = cadInstance.Document.GetElement(curve.GraphicsStyleId) as GraphicsStyle;

                if (graphicsStyle != null)
                {
                    // Get the category of the graphics style
                    Category styleCategory = graphicsStyle.GraphicsStyleCategory;

                    if (styleCategory != null)
                    {
                        // Get the name of the category, which represents the layer name in CAD
                        layerName = styleCategory.Name;
                    }
                }
            }
            // Kiểm tra xem geometryObject có phải là đường nét đa giác hay không
            else if (geometryObject is PolyLine polyline)
            {
                // Get the graphics style of the polyline
                GraphicsStyle graphicsStyle = cadInstance.Document.GetElement(polyline.GraphicsStyleId) as GraphicsStyle;

                if (graphicsStyle != null)
                {
                    // Get the category of the graphics style
                    Category styleCategory = graphicsStyle.GraphicsStyleCategory;

                    if (styleCategory != null)
                    {
                        // Get the name of the category, which represents the layer name in CAD
                        layerName = styleCategory.Name;
                    }
                }
            }

            return layerName;
        }

        //public static string GetLayerNameFromCurve(Curve curve, ImportInstance cadInstance)
        //{
        //    string layerName = "";

        //    // Get the graphics style of the curve
        //    GraphicsStyle graphicsStyle = cadInstance.Document.GetElement(curve.GraphicsStyleId) as GraphicsStyle;

        //    if (graphicsStyle != null)
        //    {
        //        // Get the category of the graphics style
        //        Category styleCategory = graphicsStyle.GraphicsStyleCategory;

        //        if (styleCategory != null)
        //        {
        //            // Get the name of the category, which represents the layer name in CAD
        //            layerName = styleCategory.Name;
        //        }
        //    }

        //    return layerName;
        //}

        public static List<Solid> GetSolids(ImportInstance cadInstance)
        {
            List<Solid> allSolids = new List<Solid>();

            // Defaults to medium detail, no references and no view.
            Options option = new Options();
            GeometryElement geoElement = cadInstance.get_Geometry(option);

            foreach (GeometryObject geoObject in geoElement)
            {
                if (geoObject is GeometryInstance)
                {
                    GeometryInstance geoInstance = geoObject as GeometryInstance;
                    GeometryElement geoElement2 = geoInstance.GetInstanceGeometry();
                    foreach (GeometryObject geoObject2 in geoElement2)
                    {
                        if (geoObject2 is Solid)
                        {
                            Solid solid = geoObject2 as Solid;
                            if (solid.SurfaceArea > 0) allSolids.Add(solid);
                        }
                    }
                }
            }

            return allSolids;
        }

        public static List<PlanarFace> GetHatchHaveName(ImportInstance cadInstance,
            string layerName)
        {
            List<PlanarFace> allHatch = new List<PlanarFace>();

            List<Solid> solids = GetSolids(cadInstance);
            if (solids.Count == 0) return allHatch;

            foreach (Solid solid in solids)
            {
                foreach (Face face in solid.Faces)
                {
                    GraphicsStyle graphicsStyle =
                           cadInstance.Document.GetElement(face.GraphicsStyleId)
                           as GraphicsStyle;

                    if (graphicsStyle == null) continue;
                    Category styleCategory = graphicsStyle.GraphicsStyleCategory;

                    if (styleCategory.Name.Equals(layerName))
                    {
                        allHatch.Add(face as PlanarFace);
                    }
                }
            }

            return allHatch;
        }

        public static List<Arc> GetArcsHaveName(ImportInstance cadInstance, string layerName)
        {
            List<Arc> allArcs = new List<Arc>();

            // Defaults to medium detail, no references and no view.
            Options option = new Options();
            GeometryElement geoElement = cadInstance.get_Geometry(option);

            foreach (GeometryObject arcObject in geoElement)
            {
                if (arcObject is GeometryInstance)
                {
                    GeometryInstance geoInstance = arcObject as GeometryInstance;
                    GeometryElement geoElement2 = geoInstance.GetInstanceGeometry();
                    foreach (GeometryObject arcObject2 in geoElement2)
                    {
                        if (arcObject2 is Arc)
                        {
                            Arc arc = arcObject2 as Arc;
                            GraphicsStyle graphicsStyle = cadInstance.Document.GetElement(arc.GraphicsStyleId) as GraphicsStyle;
                            //if (graphicsStyle == null) continue;
                            Category styleCategory = graphicsStyle.GraphicsStyleCategory;
                            if (arc.Length > 0 && styleCategory.Name.Equals(layerName))
                            {
                                allArcs.Add(arc);
                            }
                        }
                    }
                }
            }

            return allArcs;
        }
        public static List<Line> GetLinesHaveName(ImportInstance cadInstance)
        {
            List<Line> allLines = new List<Line>();

            // Defaults to medium detail, no references and no view.
            Options option = new Options();
            GeometryElement geoElement = cadInstance.get_Geometry(option);

            foreach (GeometryObject lineObject in geoElement)
            {
                if (lineObject is GeometryInstance)
                {
                    GeometryInstance geoInstance = lineObject as GeometryInstance;
                    GeometryElement geoElement2 = geoInstance.GetInstanceGeometry();
                    foreach (GeometryObject lineObject2 in geoElement2)
                    {
                        if (lineObject2 is Line)
                        {
                            Line line = lineObject2 as Line;
                            GraphicsStyle graphicsStyle = cadInstance.Document.GetElement(line.GraphicsStyleId) as GraphicsStyle;
                            Category styleCategory = graphicsStyle.GraphicsStyleCategory;
                            if (line.Length > 0)
                            {
                                allLines.Add(line);
                            }
                        }
                    }
                }
            }

            return allLines;
        }
        public static List<Curve> GetAllCurves(ImportInstance cadInstance)
        {
            List<Curve> allCurves = new List<Curve>();

            // Defaults to medium detail, no references and no view.
            Options option = new Options();
            GeometryElement geoElement = cadInstance.get_Geometry(option);

            foreach (GeometryObject geoObject in geoElement)
            {
                if (geoObject is GeometryInstance geoInstance)
                {
                    GeometryElement geoElement2 = geoInstance.GetInstanceGeometry();
                    foreach (GeometryObject geoObject2 in geoElement2)
                    {
                        if (geoObject2 is Curve curve)
                        {
                            allCurves.Add(curve);
                        }
                    }
                }
                else if (geoObject is Curve curve)
                {
                    allCurves.Add(curve);
                }
            }

            return allCurves;
        }

        public static List<Curve> GetAllLinesAndPolylines(ImportInstance importInstance)
        {
            List<Curve> allLinesAndPolylines = new List<Curve>();

            // Defaults to medium detail, no references and no view.
            Options option = new Options();
            GeometryElement geoElement = importInstance.get_Geometry(option);

            foreach (GeometryObject geoObject in geoElement)
            {
                if (geoObject is GeometryInstance geoInstance)
                {
                    GeometryElement geoElement2 = geoInstance.GetInstanceGeometry();
                    foreach (GeometryObject geoObject2 in geoElement2)
                    {
                        if (geoObject2 is Line line)
                        {
                            allLinesAndPolylines.Add(line);
                        }
                        else if (geoObject2 is PolyLine polyLine)
                        {
                            IList<XYZ> points = polyLine.GetCoordinates();
                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                Line segment = Line.CreateBound(points[i], points[i + 1]);
                                allLinesAndPolylines.Add(segment);
                            }
                        }
                    }
                }
                
            }

            return allLinesAndPolylines;
        }

        //public static List<XYZ> GetAllBlockOrigin(ImportInstance importInstance)
        //{
        //    List<XYZ> blockCenters = new List<XYZ>();
        //    GeometryElement geoElement = importInstance.get_Geometry(new Options());
        //    foreach (GeometryObject geoObject in geoElement)
        //    {
        //        if (geoObject is GeometryInstance instance)
        //        {
        //            GeometryElement geoElement2 = instance.GetInstanceGeometry();
        //            foreach (GeometryObject geoObject2 in geoElement2)
        //            {
        //                if (geoObject2 is GeometryInstance blockInstance)
        //                {
        //                    XYZ blockCenter = blockInstance.Transform.Origin;
        //                    blockCenters.Add(blockCenter);
        //                }
        //            }
        //        }
        //    }
        //    return blockCenters;
        //}

        public static List<XYZ> GetAllBlockOrigin(ImportInstance importInstance)
        {
            List<XYZ> blockCenters = new List<XYZ>();
            GeometryElement geoElement = importInstance.get_Geometry(new Options());
            foreach (GeometryObject geoObject in geoElement)
            {
                GeometryInstance instance = geoObject as GeometryInstance;
                if (instance == null) continue;
                foreach (GeometryObject geoObject2 in instance.SymbolGeometry)
                {
                    if (!(geoObject2 is GeometryInstance blockInstance)) continue;
                    XYZ blockCenter = blockInstance.Transform.Origin;
                    blockCenters.Add(blockCenter);
                }
            }

            return blockCenters;
        }

        public static List<Arc> GetArcsFromImportInstance(ImportInstance importInstance)
        {
            List<Arc> arcs = new List<Arc>();

            // Lấy phần tử hình học từ ImportInstance
            GeometryElement geoElement = importInstance.get_Geometry(new Options());

            // Duyệt qua các đối tượng hình học
            foreach (GeometryObject geoObject in geoElement)
            {
                if (geoObject is GeometryInstance instance)
                {
                    // Lặp qua các đối tượng hình học trong SymbolGeometry
                    foreach (GeometryObject geoObject2 in instance.SymbolGeometry)
                    {
                        if (geoObject2 is GeometryInstance blockInstance)
                        {
                            // Lặp qua các đối tượng hình học trong khối hình học này
                            foreach (GeometryObject geoObject3 in blockInstance.GetInstanceGeometry())
                            {
                                // Kiểm tra nếu đối tượng là Arc
                                if (geoObject3 is Arc arc)
                                {
                                    // Thêm Arc vào danh sách kết quả
                                    arcs.Add(arc);
                                }
                            }
                        }
                    }
                }
            }

            return arcs;
        }


    }
}