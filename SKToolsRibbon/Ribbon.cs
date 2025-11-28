using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;

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
            const string ribbonName = "SK-Tools";

            try { uiapp.CreateRibbonTab(ribbonName); }
            catch { /* Tab đã tồn tại */ }

            var GENPanel = uiapp.CreateRibbonPanel(ribbonName, "GEN's Addins");
            var CADPanel = uiapp.CreateRibbonPanel(ribbonName, "CAD's Addins");
            var MEPPanel = uiapp.CreateRibbonPanel(ribbonName, "MEP's Addins");

            string execLocation = Assembly.GetExecutingAssembly().Location;

            string dllDir = Path.GetDirectoryName(execLocation);

            string bundleDir = Directory.GetParent(dllDir).FullName;

            // 3) Xây lại path tới SKRevitAddins.dll
            string dllPath = Path.Combine(bundleDir, "dll", "SKRevitAddins.dll");

            void AddButton(RibbonPanel panel, string name, string text, string className, string icon, string tooltip = null)
            {
                var btn = ribbonUtils.CreatePushButtonData(name, text, dllPath, className, icon, tooltip);
                if (btn != null) panel.AddItem(btn);
            }

            // GEN Panel
            AddButton(GENPanel, "CopySetOfFilterFromViewTempCmd", "Filter\nCopy", "SKRevitAddins.CopySetOfFilterFromViewTemp.CopySetOfFilterFromViewTempCmd", "CopySetOfFilterFromViewTemp.png");
            AddButton(GENPanel, "SelectElementsVer1Cmd", "Elements\nNumbering", "SKRevitAddins.SelectElementsVer1.SelectElementsVer1Cmd", "SelectElements.png");
            AddButton(GENPanel, "ReplaceTextNoteFromCadExploreCmd", "TextNotes\nEditor", "SKRevitAddins.CadImportReplaceTextType.CadImportReplaceTextTypeCmd", "ReplaceText.png");
            AddButton(GENPanel, "ExportSchedulesToExcelCmd", "Schedules\nExport", "SKRevitAddins.ExportSchedulesToExcel.ExportSchedulesToExcelCmd", "SchedulesToExcel.png");
            AddButton(GENPanel, "RefPointToTopoCmd", "PCL\nTopo", "SKRevitAddins.PointCloudAddins.RefPointToTopo.RefPointToTopoCmd", "ScanAndModel.png");
            AddButton(GENPanel, "TogglePointCloudCmd", "PCL\nVisible", "SKRevitAddins.PointCloudAddins.TogglePointCloudCmd", "TogglePointCloud.png");
            AddButton(GENPanel, "CheckAndMapPointCloudCmd", "PCL\nMapping", "SKRevitAddins.PointCloudAddins.CheckAndMapPointCloudCmd", "CheckAndMapPointCloud.png");
            AddButton(GENPanel, "CreateSheetsFromExcelCmd", "Sheet\nCreate", "SKRevitAddins.CreateSheetsFromExcel.CreateSheetsFromExcelCmd", "CreateSheetsFromExcel.png");
            AddButton(GENPanel, "GridToolsCmd", "Grid\nTools", "SKRevitAddins.GridTools.GridToolsCmd", "GridTools.png");
            AddButton(GENPanel, "TagTools", "TAGs\nTool", "SKRevitAddins.TAGTools.TagToolsCmd", "TagAlign.png");
            AddButton(GENPanel, "FLTools", "Create FLs\nTool", "SKRevitAddins.GENTools.BatchCreateOrDeletePlansCmd", "BatchCreateFloorPlans.png");
            AddButton(GENPanel, "DelIDs", "Duplicate\nDelete", "SKRevitAddins.GENTools.DuplicateElementsDetectorCmd", "DuplicateElementsDetector.png");
            AddButton(GENPanel, "CopyInfor", "Copy\nParam", "SKRevitAddins.GENTools.CopyParamFromLinkCmd", "CopyParamFromLink.png");
            AddButton(GENPanel, "RoomNameToElements", "RoomName\n ToElements", "SKRevitAddins.GENTools.RoomDataToElementsCmd", "RoomDataToElements.png");

            // CAD Panel
            AddButton(CADPanel, "AutoPlaceElementFrBlockCADCmd", "Elements\nCreate", "SKRevitAddins.AutoPlaceElementFrBlockCAD.AutoPlaceElementFrBlockCADCmd", "PlaceEleFromBlocks.png", "Create elements from Blocks");
            AddButton(CADPanel, "FindDWGNotUseAndDelCmd", "CADFind\nDelete", "SKRevitAddins.FindDWGNotUseAndDel.FindDWGNotUseAndDelCmd", "FindDWGNotUsed.png", "Find DWG and Delete");
            AddButton(CADPanel, "AutoCreatePileFromCadCmd", "Piles\n Create", "SKRevitAddins.AutoCreatePileFromCad.AutoCreatePileFromCadCmd", "PileFCADCreate.png", "Create Pile From DWG");
            AddButton(GENPanel, "LayoutsToDWGCmd", "Layout\nExport", "SKRevitAddins.LayoutsToDWG.LayoutsToDWGCmd", "DWGExport.png");

            // MEP Panel
            AddButton(MEPPanel, "PlaceDuctsAndPipesBaseonCadCmd", "Duct・Pipes\n Create", "SKRevitAddins.DuctPipePlaceholderAndFittings.PlaceDuctsAndPipesBaseonCadCmd", "PlaceDuctsAndPipesBaseonCad.png");
            AddButton(MEPPanel, "ConvertDuctsAndPipesToPlaceholdersCmd", "Duct・Pipes\n Convert", "SKRevitAddins.DuctPipePlaceholderAndFittings.ConvertDuctsAndPipesToPlaceholdersCmd", "ConvertDuctsAndPipesToPalceholders.png");
            AddButton(MEPPanel, "ConvertPlaceholdersToDuctsAndPipesCmd", "Placeholders\n Convert", "SKRevitAddins.DuctPipePlaceholderAndFittings.ConvertPlaceholdersToDuctsAndPipesCmd", "ConvertPlaceholdersToDuctsAndPipes.png");
            AddButton(MEPPanel, "IntersectWithFrameCmd", "Permissible\n Check", "SKRevitAddins.PermissibleRangeFrame.PermissibleRangeFrameCmd", "PermissibleRangeFramePunching.png");
            AddButton(MEPPanel, "SleeveCheckerCmd", "Sleeve\n Checker", "SKRevitAddins.SleeveChecker.SleeveCheckerCmd", "PermissibleRangeFramePunching.png");
        }
    }
}
