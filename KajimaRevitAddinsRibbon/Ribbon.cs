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
            //    "スペース\n作成", "SKToolsAddins.dll",
            //    "SKToolsAddins.Commands.CreateSpace.CreateSpaceCmd", "CreateSpace.png",
            //    "アドインの情報", null, null, null, null, null);

            //createSpacePanel.AddItem(createSpaceBtn);

            //// ChangeBwTypAndInsPara
            //PushButtonData changeParaBtn = ribbonUtils.CreatePushButtonData("ChangeBwTypAndInsParaCmd",
            //    "パラメーター\n変更", "SKToolsAddins.dll",
            //    "SKToolsAddins.Commands.ChangeBwTypeAndIns.ChangeBwTypeAndInsCmd", "ChangeBwTypAndInsPara.png",
            //    "アドインの情報", null, null, null, null, null);

            //changeParaPanel.AddItem(changeParaBtn);

            // CopySetOfFilterFromViewTemp
            PushButtonData copySetOfFilterFromViewTempBtn = ribbonUtils.CreatePushButtonData("CopySetOfFilterFromViewTempCmd",
                "Filter\nCopy", "SKToolsAddins.dll",
                "SKToolsAddins.Commands.CopySetOfFilterFromViewTemp.CopySetOfFilterFromViewTempCmd", "CopySetOfFilterFromViewTemp.png",
                "アドインの情報", null, null, null, null, null);

            GENPanelNamePanel.AddItem(copySetOfFilterFromViewTempBtn);

            // SelectElements
            PushButtonData selectElementsBtn = ribbonUtils.CreatePushButtonData("SelectElementsVer1Cmd"
                , "Elements\nNumbering"
                , "SKToolsAddins.dll"
                , "SKToolsAddins.Commands.SelectElementsVer1.SelectElementsVer1Cmd"
                , "SelectElements.png"
                , "アドインの情報", null, null, null, null, null);
            GENPanelNamePanel.AddItem(selectElementsBtn);

            // Parameter Assignment from excel for Jasty
            PushButtonData parameterAssigmentBtn = ribbonUtils.CreatePushButtonData("ParameterAssignmentCmd"
                , "パラメーター\n作成"
                , "SKToolsAddins.dll"
                , "SKToolsAddins.Commands.ParameterAssignment.ParameterAssignmentCmd"
                , "SelectElements.png"
                , "", null, null, null, null, null);
            GENPanelNamePanel.AddItem(parameterAssigmentBtn);

            //CadImportReplaceTextType
            PushButtonData replaceTextNotesBtn = ribbonUtils.CreatePushButtonData("ReplaceTextNoteFromCadExploreCmd",
                "TextType\nReplace", "SKToolsAddins.dll",
                "SKToolsAddins.Commands.CadImportReplaceTextType.CadImportReplaceTextTypeCmd", "ReplaceText.png",
                "アドインの情報", null, null, null, null, null);
            CADPanel.AddItem(replaceTextNotesBtn);

            //CadImportReplaceLineStyle
            PushButtonData replaceLineStyleBtn = ribbonUtils.CreatePushButtonData("ReplaceLineStyleFromCadExploreCmd",
                "LineStyle\nReplace", "SKToolsAddins.dll",
                "SKToolsAddins.Commands.CadImportReplaceLineStyle.CadImportReplaceLineStyleCmd", "ReplaceLine.png",
                "アドインの情報", null, null, null, null, null);

            CADPanel.AddItem(replaceLineStyleBtn);

            //CadLinkPlaceElementsFromBlocks
            PushButtonData placeEleFromBlocksBtn = ribbonUtils.CreatePushButtonData("PlaceElementsFromBlocksCadCmd",
                "Elements\nPlace", "SKToolsAddins.dll",
                "SKToolsAddins.Commands.PlaceElementsFromBlocksCad.PlaceElementsFromBlocksCadCmd", "ReplaceText.png",
                "アドインの情報", null, null, null, null, null);
            CADPanel.AddItem(placeEleFromBlocksBtn);

            // FindDWGNotUsed
            PushButtonData findDWGNotUsedBtn = ribbonUtils.CreatePushButtonData("FindDWGNotUsedAndDelCmd"
                , "Import DWG\nDelete"
                , "SKToolsAddins.dll",
                "SKToolsAddins.Commands.FindDWGNotUsedAndDel.FindDWGNotUsedAndDelCmd"
                , "FindDWGNotUsed.png"
                , "アドインの情報"
                , null, null, null, null, null);
            CADPanel.AddItem(findDWGNotUsedBtn);

            // Create Pile From Cad
            PushButtonData createPileFromCadBtn = ribbonUtils.CreatePushButtonData("AutoCreatePileFromCadCmd"
                , "Piles DWG\n Import"
                , "SKToolsAddins.dll",
                "SKToolsAddins.Commands.AutoCreatePileFromCad.AutoCreatePileFromCadCmd"
                , "ReplaceLine.png"
                , "アドインの情報"
                , null, null, null, null, null);
            CADPanel.AddItem(createPileFromCadBtn);

            // Create Duct/ Pipe from CAD

            PushButtonData createDuctAndPipeFromCadBtn = ribbonUtils.CreatePushButtonData("PlaceDuctsAndPipesBaseonCadCmd"
                , "Duct・Pipes\n Create"
                , "SKToolsAddins.dll",
                "SKToolsAddins.Commands.DuctPipePlaceholderAndFittings.PlaceDuctsAndPipesBaseonCadCmd"
                , "PlaceDuctsAndPipesBaseonCad.png"
                , null
                , null, null, null, null, null);
            mepAddinsPanel.AddItem(createDuctAndPipeFromCadBtn);

            PushButtonData convertDuctsAndPipesToPalceholdersBtn = ribbonUtils.CreatePushButtonData("ConvertDuctsAndPipesToPlaceholdersCmd"
                , "Duct・Pipes\n Convert"
                , "SKToolsAddins.dll",
                "SKToolsAddins.Commands.DuctPipePlaceholderAndFittings.ConvertDuctsAndPipesToPlaceholdersCmd"
                , "ConvertDuctsAndPipesToPalceholders.png"
                , null
                , null, null, null, null, null);
            mepAddinsPanel.AddItem(convertDuctsAndPipesToPalceholdersBtn);

            PushButtonData ConvertPlaceholdersToDuctsAndPipesBtn = ribbonUtils.CreatePushButtonData("ConvertPlaceholdersToDuctsAndPipesCmd"
                , "Placeholders\n Convert"
                , "SKToolsAddins.dll",
                "SKToolsAddins.Commands.DuctPipePlaceholderAndFittings.ConvertPlaceholdersToDuctsAndPipesCmd"
                , "ConvertPlaceholdersToDuctsAndPipes.png"
                , null
                , null, null, null, null, null);
            mepAddinsPanel.AddItem(ConvertPlaceholdersToDuctsAndPipesBtn);
        }
    }
}
