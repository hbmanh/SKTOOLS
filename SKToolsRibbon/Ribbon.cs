using Autodesk.Revit.UI;
using RibbonPanel = Autodesk.Revit.UI.RibbonPanel;

namespace SKToolsRibbon
{
    public class Ribbon : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application) => Result.Succeeded;

        public Result OnStartup(UIControlledApplication application)
        {
            InitializeRibbon(application);
            return Result.Succeeded;
        }

        private void InitializeRibbon(UIControlledApplication uiapp)
        {
            var ribbonUtils = new RibbonUtils(uiapp.ControlledApplication);
            string ribbonName = "SK-Tools";
            uiapp.CreateRibbonTab(ribbonName);

            var GENPanel = uiapp.CreateRibbonPanel(ribbonName, "GEN's Addins");
            var CADPanel = uiapp.CreateRibbonPanel(ribbonName, "CAD's Addins");
            var MEPPanel = uiapp.CreateRibbonPanel(ribbonName, "MEP's Addins");

            // Helper method
            void AddButton(RibbonPanel panel, string name, string text, string className, string icon, string tooltip = null)
            {
                var btn = ribbonUtils.CreatePushButtonData(name, text, "SKRevitAddins.dll", className, icon, tooltip, null, null, null, null, null);
                if (btn != null) panel.AddItem(btn);
            }

            // GEN Panel
            AddButton(GENPanel, 
                "CopySetOfFilterFromViewTempCmd", "Filter\nCopy", "SKRevitAddins.Commands.CopySetOfFilterFromViewTemp.CopySetOfFilterFromViewTempCmd", "CopySetOfFilterFromViewTemp.png");
            AddButton(GENPanel, 
                "SelectElementsVer1Cmd", "Elements\nNumbering", "SKRevitAddins.Commands.SelectElementsVer1.SelectElementsVer1Cmd", "SelectElements.png");
            AddButton(GENPanel, 
                "ReplaceTextNoteFromCadExploreCmd", "TextNotes\nEditor", "SKRevitAddins.Commands.CadImportReplaceTextType.CadImportReplaceTextTypeCmd", "ReplaceText.png");
            AddButton(GENPanel, 
                "ExportSchedulesToExcelCmd", "Schedules\nExport", "SKRevitAddins.Commands.ExportSchedulesToExcel.ExportSchedulesToExcelCmd", "SchedulesToExcel.png");
            AddButton(GENPanel, 
                "ScanAndModelCmd", "PCL\nModel", "SKRevitAddins.Commands.ScanAndModel.ScanAndModelCmd", "ScanAndModel.png");
            AddButton(GENPanel, 
                "CreateSheetsFromExcelCmd", "Sheet\nCreate", "SKRevitAddins.Commands.CreateSheetsFromExcel.CreateSheetsFromExcelCmd", "CreateSheetsFromExcel.png");
            AddButton(GENPanel,
                "GridToolsCmd", "Grid\nTools", "SKRevitAddins.Commands.GridTools.GridToolsCmd", "GridTools.png");
            // CAD Panel
            AddButton(CADPanel, 
                "PlaceElementsFromBlocksCadCmd", "EleFrB\nCreate", "SKRevitAddins.Commands.PlaceElementsFromBlocksCad.PlaceElementsFromBlocksCadCmd", "PlaceEleFromBlocks.png", "Creeate elements from Blocks");
            AddButton(CADPanel, 
                "FindDWGNotUseAndDelCmd", "CADFind\nDelete", "SKRevitAddins.Commands.FindDWGNotUseAndDel.FindDWGNotUseAndDelCmd", "FindDWGNotUsed.png", "Find DWG and Delete");
            AddButton(CADPanel, 
                "AutoCreatePileFromCadCmd", "Piles\n Create", "SKRevitAddins.Commands.AutoCreatePileFromCad.AutoCreatePileFromCadCmd", "PileFCADCreate.png", "Create Pile From DWG");
            AddButton(GENPanel,
                "DWGExportCmd", "DWG\nExport", "SKRevitAddins.Commands.DWGExport.DWGExportCmd", "DWGExport.png");
            AddButton(GENPanel,
                "LayoutsToDWGCmd", "Layout\nExport", "SKRevitAddins.Commands.LayoutsToDWG.LayoutsToDWGCmd", "DWGExport.png");
            // MEP Panel
            AddButton(MEPPanel, 
                "PlaceDuctsAndPipesBaseonCadCmd", "Duct・Pipes\n Create", "SKRevitAddins.Commands.DuctPipePlaceholderAndFittings.PlaceDuctsAndPipesBaseonCadCmd", "PlaceDuctsAndPipesBaseonCad.png", "Create Duct/ Pipes from DWG");
            AddButton(MEPPanel, 
                "ConvertDuctsAndPipesToPlaceholdersCmd", "Duct・Pipes\n Convert", "SKRevitAddins.Commands.DuctPipePlaceholderAndFittings.ConvertDuctsAndPipesToPlaceholdersCmd", "ConvertDuctsAndPipesToPalceholders.png");
            AddButton(MEPPanel, 
                "ConvertPlaceholdersToDuctsAndPipesCmd", "Placeholders\n Convert", "SKRevitAddins.Commands.DuctPipePlaceholderAndFittings.ConvertPlaceholdersToDuctsAndPipesCmd", "ConvertPlaceholdersToDuctsAndPipes.png");
            AddButton(MEPPanel, 
                "IntersectWithFrameCmd", "Permissible\n Check", "SKRevitAddins.Commands.PermissibleRangeFrame.PermissibleRangeFrameCmd", "PermissibleRangeFramePunching.png");
        }
    }
}
