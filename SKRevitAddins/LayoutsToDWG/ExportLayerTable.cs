using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace SKRevitAddins.Utils
{
    public class ExportLayerTable
    {
        public class LayerRow
        {
            public string CategoryName { get; set; }
            public string SubCategoryName { get; set; }
            public string ProjectionLayer { get; set; }
            public int ProjectionColorId { get; set; }
            public string CutLayer { get; set; }
            public int CutColorId { get; set; }
        }

        public static List<LayerRow> GetExportLayerTable(Document doc)
        {
            var result = new List<LayerRow>();

            Categories categories = doc.Settings.Categories;
            foreach (Category cat in categories)
            {
                if (!cat.IsVisibleInUI) continue;

                // Projection & Cut info for main category
                var layerRow = new LayerRow
                {
                    CategoryName = cat.Name,
                    SubCategoryName = "",
                    ProjectionLayer = cat.GetGraphicsStyle(GraphicsStyleType.Projection)?.GraphicsStyleCategory?.Name ?? "",
                    ProjectionColorId = GetColorSafe(cat.LineColor),
                    CutLayer = cat.GetGraphicsStyle(GraphicsStyleType.Cut)?.GraphicsStyleCategory?.Name ?? "",
                    CutColorId = GetColorSafe(cat.LineColor)
                };
                result.Add(layerRow);

                // Subcategories
                foreach (Category subCat in cat.SubCategories)
                {
                    if (subCat == null || !subCat.IsVisibleInUI) continue;

                    var subLayerRow = new LayerRow
                    {
                        CategoryName = cat.Name,
                        SubCategoryName = subCat.Name,
                        ProjectionLayer = subCat.GetGraphicsStyle(GraphicsStyleType.Projection)?.GraphicsStyleCategory?.Name ?? "",
                        ProjectionColorId = GetColorSafe(subCat.LineColor),
                        CutLayer = subCat.GetGraphicsStyle(GraphicsStyleType.Cut)?.GraphicsStyleCategory?.Name ?? "",
                        CutColorId = GetColorSafe(subCat.LineColor)
                    };
                    result.Add(subLayerRow);
                }
            }

            return result;
        }

        private static int GetColorSafe(Color color)
        {
            return color != null ? color.Red : 0;
        }
    }
}
