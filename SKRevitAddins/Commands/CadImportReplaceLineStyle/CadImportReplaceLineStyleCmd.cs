using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRevitAddins.Commands.CadImportReplaceLineStyle
{
    [Transaction(TransactionMode.Manual)]
    public class CadImportReplaceLineStyleCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var lineCollector = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfCategory(BuiltInCategory.OST_Lines)
                .WhereElementIsNotElementType()
                .Cast<CurveElement>();

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Create and Replace Line Styles");

                foreach (var line in lineCollector)
                {
                    // Get Line Pattern, Line Weight, and Color
                    string linePatternName;
                    var lineStyleId = line.LineStyle.Id;
                    var graphicsStyle = doc.GetElement(lineStyleId) as GraphicsStyle;
                    if (graphicsStyle == null) continue;
                    int lineWeight = graphicsStyle.GraphicsStyleCategory.GetLineWeight(GraphicsStyleType.Projection).Value;

                    var linePatternId = graphicsStyle.GraphicsStyleCategory.GetLinePatternId(GraphicsStyleType.Projection);
                    LinePatternElement linePatternElement = doc.GetElement(linePatternId) as LinePatternElement;
                    Color color = graphicsStyle.GraphicsStyleCategory.LineColor;
                    linePatternName = linePatternElement == null ? "Solid" : linePatternElement.Name;

                    // Create Line Style Name based on Line Pattern, Line Weight, and Color
                    string newLineStyleName = $"{linePatternName}_{lineWeight}_{color.Red}-{color.Green}-{color.Blue}";
                    // Add the new linestyle 
                    Categories categories = doc.Settings.Categories;
                    Category lineCat = categories.get_Item(BuiltInCategory.OST_Lines);

                    // Check if Line Style Name already exists in the project
                    Category newLineStyleCat = doc.Settings.Categories
                        .get_Item(BuiltInCategory.OST_Lines)
                        .SubCategories.Cast<Category>() // Convert to IEnumerable<Category>
                        .FirstOrDefault(x => x.Name == newLineStyleName);

                    if (newLineStyleCat == null)
                    {
                        // If Line Style Name does not exist, create new Line Style
                        newLineStyleCat = categories.NewSubcategory(lineCat, newLineStyleName);
                        newLineStyleCat.SetLineWeight(lineWeight, GraphicsStyleType.Projection);
                        newLineStyleCat.LineColor = color;
                        newLineStyleCat.SetLinePatternId(linePatternId, GraphicsStyleType.Projection);
                    }
                    // Add to dictionary for future reference
                    doc.Regenerate();

                    Category LinesCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                    Category newLineStyleSubCate = LinesCat.SubCategories.get_Item(newLineStyleName);
                    GraphicsStyle newLineStyle = newLineStyleSubCate.GetGraphicsStyle(GraphicsStyleType.Projection);
                    line.LineStyle = newLineStyle;
                }
                trans.Commit();
            }
            TaskDialog.Show("Success", "Successfully applied new LineStyle.");

            return Result.Succeeded;
        }
    }
}
