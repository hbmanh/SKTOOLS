using Autodesk.Revit.UI;
using RibbonPanel = Autodesk.Revit.UI.RibbonPanel;

namespace SKToolsRibbon
{
    public class Ribbon : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            InitializeRibbon(application);
            return Result.Succeeded;
        }

        private void InitializeRibbon(UIControlledApplication uiapp)
        {
            RibbonUtils ribbonUtils = new RibbonUtils(uiapp.ControlledApplication);
            RibbonConstraints ribbonConstraints = new RibbonConstraints(uiapp.ControlledApplication);

            // Create Ribbon Tab
            string ribbonName = "SK-Tools";
            uiapp.CreateRibbonTab(ribbonName);

            //string createSpacePanelName = "スペースー括作成";
            //RibbonPanel createSpacePanel = uiapp.CreateRibbonPanel(ribbonName, createSpacePanelName);

            //string changeParaPanelName = "パラメーター変更";
            //RibbonPanel changeParaPanel = uiapp.CreateRibbonPanel(ribbonName, changeParaPanelName);

            string GENPanelName = "GEN's Addins";
            RibbonPanel GENPanelNamePanel = uiapp.CreateRibbonPanel(ribbonName, GENPanelName);

            //string selectElementsPanelName = "Select Elements";
            //RibbonPanel selectElementsPanel = uiapp.CreateRibbonPanel(ribbonName, selectElementsPanelName);

            string CADPanelName = "CAD's Addins";
            RibbonPanel CADPanel = uiapp.CreateRibbonPanel(ribbonName, CADPanelName);

            string mepAddinsPanelName = "MEP's Addins";
            RibbonPanel mepAddinsPanel = uiapp.CreateRibbonPanel(ribbonName, mepAddinsPanelName);

            // SK'sTools Panel



            // Create Space
            //PushButtonData createSpaceBtn = ribbonUtils.CreatePushButtonData("CreateSpaceCmd",
            //    "スペース\n作成", "SKRevitAddins.dll",
            //    "SKRevitAddins.Commands.CreateSpace.CreateSpaceCmd", "CreateSpace.png",
            //    "アドインの情報", null, null, null, null, null);

            //createSpacePanel.AddItem(createSpaceBtn);

            //// ChangeBwTypAndInsPara
            //PushButtonData changeParaBtn = ribbonUtils.CreatePushButtonData("ChangeBwTypAndInsParaCmd",
            //    "パラメーター\n変更", "SKRevitAddins.dll",
            //    "SKRevitAddins.Commands.ChangeBwTypeAndIns.ChangeBwTypeAndInsCmd", "ChangeBwTypAndInsPara.png",
            //    "アドインの情報", null, null, null, null, null);

            //changeParaPanel.AddItem(changeParaBtn);

            // CopySetOfFilterFromViewTemp
            PushButtonData copySetOfFilterFromViewTempBtn = ribbonUtils.CreatePushButtonData("CopySetOfFilterFromViewTempCmd",
                "Filter\nCopy", "SKRevitAddins.dll",
                "SKRevitAddins.Commands.CopySetOfFilterFromViewTemp.CopySetOfFilterFromViewTempCmd", "CopySetOfFilterFromViewTemp.png",
                "アドインの情報", null, null, null, null, null);

            GENPanelNamePanel.AddItem(copySetOfFilterFromViewTempBtn);

            // SelectElements
            PushButtonData selectElementsBtn = ribbonUtils.CreatePushButtonData("SelectElementsVer1Cmd"
                , "Elements\nNumbering"
                , "SKRevitAddins.dll"
                , "SKRevitAddins.Commands.SelectElementsVer1.SelectElementsVer1Cmd"
                , "SelectElements.png"
                , "アドインの情報", null, null, null, null, null);
            GENPanelNamePanel.AddItem(selectElementsBtn);

            //// Parameter Assignment from excel for Jasty
            //PushButtonData parameterAssigmentBtn = ribbonUtils.CreatePushButtonData("ParameterAssignmentCmd"
            //    , "パラメーター\n作成"
            //    , "SKRevitAddins.dll"
            //    , "SKRevitAddins.Commands.ParameterAssignment.ParameterAssignmentCmd"
            //    , "SelectElements.png"
            //    , "", null, null, null, null, null);
            //GENPanelNamePanel.AddItem(parameterAssigmentBtn);

            //CadImportReplaceTextType
            PushButtonData replaceTextNotesBtn = ribbonUtils.CreatePushButtonData("ReplaceTextNoteFromCadExploreCmd",
                "TextType\nChange", "SKRevitAddins.dll",
                "SKRevitAddins.Commands.CadImportReplaceTextType.CadImportReplaceTextTypeCmd", "ReplaceText.png",
                "アドインの情報", null, null, null, null, null);
            GENPanelNamePanel.AddItem(replaceTextNotesBtn);

            ////CadImportReplaceLineStyle
            //PushButtonData replaceLineStyleBtn = ribbonUtils.CreatePushButtonData("ReplaceLineStyleFromCadExploreCmd",
            //    "LineStyle\nReplace", "SKRevitAddins.dll",
            //    "SKRevitAddins.Commands.CadImportReplaceLineStyle.CadImportReplaceLineStyleCmd", "ReplaceLine.png",
            //    "アドインの情報", null, null, null, null, null);

            //CADPanel.AddItem(replaceLineStyleBtn);

            //ExportSchedulesToExcel
            PushButtonData exportSchedulesToExcelBtn = ribbonUtils.CreatePushButtonData("ExportSchedulesToExcelCmd",
                "Schedules\nExport", "SKRevitAddins.dll",
                "SKRevitAddins.Commands.ExportSchedulesToExcel.ExportSchedulesToExcelCmd", "SchedulesToExcel.png",
                "アドインの情報", null, null, null, null, null);
            GENPanelNamePanel.AddItem(exportSchedulesToExcelBtn);

            //ScanAndModel
            PushButtonData scanAndModelBtn = ribbonUtils.CreatePushButtonData("ScanAndModelCmd",
                "PCL\nModel", "SKRevitAddins.dll",
                "SKRevitAddins.Commands.ScanAndModel.ScanAndModelCmd", "ScanAndModel.png",
                "アドインの情報", null, null, null, null, null);
            GENPanelNamePanel.AddItem(scanAndModelBtn);

            //CadLinkPlaceElementsFromBlocks
            PushButtonData placeEleFromBlocksBtn = ribbonUtils.CreatePushButtonData("PlaceElementsFromBlocksCadCmd",
                "EleFrB\nCreate", "SKRevitAddins.dll",
                "SKRevitAddins.Commands.PlaceElementsFromBlocksCad.PlaceElementsFromBlocksCadCmd", "PlaceEleFromBlocks.png",
                "Creeate elements from Blocks", null, null, null, null, null);
            CADPanel.AddItem(placeEleFromBlocksBtn);

            // FindDWGNotUsed
            PushButtonData findDWGNotUsedBtn = ribbonUtils.CreatePushButtonData("FindDWGNotUseAndDelCmd"
                , "CADFind\nDelete"
                , "SKRevitAddins.dll",
                "SKRevitAddins.Commands.FindDWGNotUseAndDel.FindDWGNotUseAndDelCmd"
                , "FindDWGNotUsed.png"
                , "Find DWG and Delete"
                , null, null, null, null, null);
            CADPanel.AddItem(findDWGNotUsedBtn);

            // Create Pile From Cad
            PushButtonData createPileFromCadBtn = ribbonUtils.CreatePushButtonData("AutoCreatePileFromCadCmd"
                , "Piles\n Create"
                , "SKRevitAddins.dll",
                "SKRevitAddins.Commands.AutoCreatePileFromCad.AutoCreatePileFromCadCmd"
                , "PileFCADCreate.png"
                , "Create Pile From DWG"
                , null, null, null, null, null);
            CADPanel.AddItem(createPileFromCadBtn);

            // Create Duct/ Pipe from CAD

            PushButtonData createDuctAndPipeFromCadBtn = ribbonUtils.CreatePushButtonData("PlaceDuctsAndPipesBaseonCadCmd"
                , "Duct・Pipes\n Create"
                , "SKRevitAddins.dll",
                "SKRevitAddins.Commands.DuctPipePlaceholderAndFittings.PlaceDuctsAndPipesBaseonCadCmd"
                , "PlaceDuctsAndPipesBaseonCad.png"
                , "Create Duct/ Pipes from DWG"
                , null, null, null, null, null);
            mepAddinsPanel.AddItem(createDuctAndPipeFromCadBtn);

            PushButtonData convertDuctsAndPipesToPalceholdersBtn = ribbonUtils.CreatePushButtonData("ConvertDuctsAndPipesToPlaceholdersCmd"
                , "Duct・Pipes\n Convert"
                , "SKRevitAddins.dll",
                "SKRevitAddins.Commands.DuctPipePlaceholderAndFittings.ConvertDuctsAndPipesToPlaceholdersCmd"
                , "ConvertDuctsAndPipesToPalceholders.png"
                , null
                , null, null, null, null, null);
            mepAddinsPanel.AddItem(convertDuctsAndPipesToPalceholdersBtn);

            PushButtonData ConvertPlaceholdersToDuctsAndPipesBtn = ribbonUtils.CreatePushButtonData("ConvertPlaceholdersToDuctsAndPipesCmd"
                , "Placeholders\n Convert"
                , "SKRevitAddins.dll",
                "SKRevitAddins.Commands.DuctPipePlaceholderAndFittings.ConvertPlaceholdersToDuctsAndPipesCmd"
                , "ConvertPlaceholdersToDuctsAndPipes.png"
                , null
                , null, null, null, null, null);
            mepAddinsPanel.AddItem(ConvertPlaceholdersToDuctsAndPipesBtn);

            PushButtonData PermissibleRangeFramePunchingBtn = ribbonUtils.CreatePushButtonData("IntersectWithFrameCmd"
                , "Permissible\n Check"
                , "SKRevitAddins.dll",
                "SKRevitAddins.Commands.PermissibleRangeFrame.PermissibleRangeFrameCmd"
                , "PermissibleRangeFramePunching.png"
                , null
                , null, null, null, null, null);
            mepAddinsPanel.AddItem(PermissibleRangeFramePunchingBtn);
        }
    }
}
