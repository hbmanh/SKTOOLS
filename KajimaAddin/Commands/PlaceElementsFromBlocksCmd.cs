//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Windows.Forms;
//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.DB.Structure;
//using Autodesk.Revit.UI;
//using Autodesk.Revit.UI.Selection;
//using ComboBox = System.Windows.Forms.ComboBox;
//using Form = System.Windows.Forms.Form;

//namespace SKToolsAddins.Commands
//{
//    [Transaction(TransactionMode.Manual)]
//    public class PlaceElementsFromBlocksCmd : IExternalCommand
//    {
//        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//        {
//            UIApplication uiapp = commandData.Application;
//            UIDocument uidoc = uiapp.ActiveUIDocument;
//            Document doc = uidoc.Document;

//            // Get the selected CAD link
//            var refLinkCad = uidoc.Selection.PickObject(ObjectType.Element, new ImportInstanceSelectionFilter(), "Select Link File");
//            var selectedCadLink = doc.GetElement(refLinkCad) as ImportInstance;
//            Level level = uidoc.ActiveView.GenLevel;

//            double offset = 0 / 304.8;

//            // Retrieve block names from the CAD link
//            List<string> blockNames = new List<string>();
//            GeometryElement geoElem = selectedCadLink.get_Geometry(new Options());
//            foreach (GeometryObject geoObj in geoElem)
//            {
//                GeometryInstance instance = geoObj as GeometryInstance;
//                if (instance == null) continue;
//                foreach (GeometryObject instObj in instance.SymbolGeometry)
//                {
//                    if (instObj is GeometryInstance blockInstance)
//                    {
//                        blockNames.Add(blockInstance.Symbol.Name);
//                    }
//                }
//            }

//            // Retrieve family names from Revit
//            List<string> familyNames = new FilteredElementCollector(doc)
//                .OfClass(typeof(FamilySymbol))
//                .Cast<FamilySymbol>()
//                .Select(fs => fs.Family.Name)
//                .Distinct()
//                .ToList();

//            // Show the form and get the user input
//            MappingForm form = new MappingForm(doc, blockNames, familyNames);
//            if (form.ShowDialog() != DialogResult.OK)
//            {
//                return Result.Cancelled;
//            }

//            var selectedCategory = (BuiltInCategory)Enum.Parse(typeof(BuiltInCategory), form.CategoryComboBox.SelectedItem.ToString());

//            Dictionary<string, (BuiltInCategory category, string familyName, string typeName)> blockMappings = new Dictionary<string, (BuiltInCategory category, string familyName, string typeName)>();
//            foreach (DataGridViewRow row in form.MappingGridView.Rows)
//            {
//                if (row.IsNewRow) continue;
//                string blockName = row.Cells[0].Value?.ToString();
//                string familyName = row.Cells[1].Value?.ToString();
//                string typeName = row.Cells[2].Value?.ToString();
//                if (!string.IsNullOrEmpty(blockName) && !string.IsNullOrEmpty(familyName) && !string.IsNullOrEmpty(typeName))
//                {
//                    blockMappings[blockName] = (selectedCategory, familyName, typeName);
//                }
//            }

//            using (Transaction trans = new Transaction(doc, "Place Elements from CAD Blocks"))
//            {
//                trans.Start();

//                // Get all block references from the CAD link
//                foreach (GeometryObject geoObj in geoElem)
//                {
//                    GeometryInstance instance = geoObj as GeometryInstance;
//                    if (instance == null) continue;
//                    foreach (GeometryObject instObj in instance.SymbolGeometry)
//                    {
//                        if (instObj is GeometryInstance blockInstance)
//                        {
//                            var blockName = blockInstance.Symbol.Name;

//                            if (blockMappings.TryGetValue(blockName, out var blockInfo))
//                            {
//                                var (category, familyName, typeName) = blockInfo;

//                                var familySymbol = new FilteredElementCollector(doc)
//                                    .OfCategory(category)
//                                    .OfClass(typeof(FamilySymbol))
//                                    .FirstOrDefault(e => (e as FamilySymbol).Family.Name == familyName && (e as FamilySymbol).Name == typeName) as FamilySymbol;

//                                if (familySymbol == null) continue;
//                                if (!familySymbol.IsActive)
//                                {
//                                    familySymbol.Activate();
//                                    doc.Regenerate();
//                                }

//                                var blockPosition = blockInstance.Transform.Origin;
//                                var blockRotation = blockInstance.Transform.BasisX.AngleTo(new XYZ(1, 0, 0));

//                                XYZ placementPosition = new XYZ(blockPosition.X, blockPosition.Y, offset);
//                                FamilyInstance familyInstance = doc.Create.NewFamilyInstance(placementPosition, familySymbol, level, StructuralType.NonStructural);

//                                // Apply the rotation
//                                ElementTransformUtils.RotateElement(doc, familyInstance.Id, Line.CreateBound(placementPosition, placementPosition + XYZ.BasisZ), blockRotation);
//                            }
//                        }
//                    }
//                }

//                trans.Commit();
//            }

//            TaskDialog.Show("Success", "Successfully placed elements at block positions from the imported CAD file.");
//            return Result.Succeeded;
//        }

//        class ImportInstanceSelectionFilter : ISelectionFilter
//        {
//            public bool AllowElement(Element elem)
//            {
//                return elem is ImportInstance;
//            }

//            public bool AllowReference(Reference reference, XYZ position)
//            {
//                return false;
//            }
//        }
//    }

//    public class MappingForm : Form
//    {
//        public ComboBox CategoryComboBox { get; private set; }
//        public DataGridView MappingGridView { get; private set; }
//        private Document _doc;

//        public MappingForm(Document doc, List<string> blockNames, List<string> familyNames)
//        {
//            _doc = doc;
//            this.Text = "Blocks to Family Mapping";
//            this.Width = 800;
//            this.Height = 500;

//            CategoryComboBox = new ComboBox
//            {
//                Left = 10,
//                Top = 10,
//                Width = 300
//            };
//            CategoryComboBox.Items.AddRange(Enum.GetNames(typeof(BuiltInCategory)).ToArray());

//            MappingGridView = new DataGridView
//            {
//                Left = 10,
//                Top = 50,
//                Width = 760,
//                Height = 380,
//                ColumnCount = 3
//            };
//            MappingGridView.Columns[0].Name = "Blocks Name";
//            MappingGridView.Columns[0].ReadOnly = true;
//            MappingGridView.Columns[1] = new DataGridViewComboBoxColumn()
//            {
//                Name = "Family Name",
//                DataSource = familyNames
//            };
//            MappingGridView.Columns[2] = new DataGridViewComboBoxColumn()
//            {
//                Name = "Type Name"
//            };

//            foreach (var blockName in blockNames)
//            {
//                MappingGridView.Rows.Add(blockName);
//            }

//            this.Controls.Add(CategoryComboBox);
//            this.Controls.Add(MappingGridView);

//            var submitButton = new Button
//            {
//                Text = "Submit",
//                Left = 350,
//                Width = 100,
//                Top = 440
//            };
//            submitButton.Click += (sender, e) => { this.DialogResult = DialogResult.OK; this.Close(); };

//            this.Controls.Add(submitButton);

//            MappingGridView.CellValueChanged += MappingGridView_CellValueChanged;
//        }

//        private void MappingGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
//        {
//            if (e.ColumnIndex == 1)
//            {
//                var familyName = MappingGridView.Rows[e.RowIndex].Cells[1].Value?.ToString();
//                if (!string.IsNullOrEmpty(familyName))
//                {
//                    var typeNames = new FilteredElementCollector(_doc)
//                        .OfClass(typeof(FamilySymbol))
//                        .Where(fs => ((FamilySymbol)fs).Family.Name == familyName)
//                        .Select(fs => ((FamilySymbol)fs).Name)
//                        .ToList();
//                    var typeNameCell = MappingGridView.Rows[e.RowIndex].Cells[2] as DataGridViewComboBoxCell;
//                    typeNameCell.DataSource = typeNames;
//                }
//            }
//        }
//    }
//}
