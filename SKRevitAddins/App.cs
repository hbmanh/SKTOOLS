using Autodesk.Revit.UI;
using SKRevitAddins.AutoCreatePileFromCad;
using SKRevitAddins.CopySetOfFilterFromViewTemp;
using SKRevitAddins.CreateSpace;
using SKRevitAddins.DeleteTypeOfTextNotesDontUse;
using SKRevitAddins.ExportSchedulesToExcel;
using SKRevitAddins.FindDWGNotUseAndDel;
using SKRevitAddins.PermissibleRangeFrame;
using SKRevitAddins.PlaceElementsFromBlocksCad;
using SKRevitAddins.SelectElements;
using SKRevitAddins.SelectElementsVer1;
using SKRevitAddins.ExportSchedulesToExcel;
using SKRevitAddins.Forms;
using SKRevitAddins.ViewModel;

namespace SKRevitAddins
{
    public class App : IExternalApplication
    {

        internal static App thisApp = null;
        private CreateSpaceWpfWindow m_CreateSpaceWpfWindow;
        private CopySetOfFilterFromViewTempWpfWindow m_CopySetOfFilterFromViewTempWpfWindow;
        private SelectElementsWpfWindow m_SelectElementsWpfWindow;
        private SelectElementsVer1WpfWindow m_SelectElementsVer1WpfWindow;
        private FindDWGNotUseAndDelWpfWindow m_FindDWGNotUseAndDelWpfWindow;
        private DeleteTypeOfTextNotesDontUseWpfWindow m_DeleteTypeOfTextNotesDontUseWpfWindow;
        private AutoCreatePileFromCadWpfWindow m_AutoCreatePileFromCadWpfWindow;
        private PlaceElementsFromBlocksCadWpfWindow m_PlaceElementsFromBlocksCadWpfWindow;
        private PermissibleRangeFrameWpfWindow m_PermissibleRangeFrameWpfWindow;
        private ExportSchedulesToExcelWpfWindow m_ExportSchedulesToExcelWpfWindow;


        public Result OnShutdown(UIControlledApplication application)
        {
            if (m_CreateSpaceWpfWindow != null && m_CreateSpaceWpfWindow.IsVisible)
            {
                m_CreateSpaceWpfWindow.Close();
            }

            if (m_CopySetOfFilterFromViewTempWpfWindow != null && m_CopySetOfFilterFromViewTempWpfWindow.IsVisible)
            {
                m_CopySetOfFilterFromViewTempWpfWindow.Close();
            }

            if (m_SelectElementsWpfWindow != null && m_SelectElementsWpfWindow.IsVisible)
            {
               m_SelectElementsWpfWindow.Close(); 
            }

            if (m_SelectElementsVer1WpfWindow != null && m_SelectElementsVer1WpfWindow.IsVisible)
            {
                m_SelectElementsVer1WpfWindow.Close();
            }
            if (m_ExportSchedulesToExcelWpfWindow != null && m_ExportSchedulesToExcelWpfWindow.IsVisible)
            {
                m_ExportSchedulesToExcelWpfWindow.Close();
            }
            if (m_FindDWGNotUseAndDelWpfWindow != null && m_FindDWGNotUseAndDelWpfWindow.IsVisible)
            {
                m_FindDWGNotUseAndDelWpfWindow.Close();
            }

            if (m_DeleteTypeOfTextNotesDontUseWpfWindow != null && m_DeleteTypeOfTextNotesDontUseWpfWindow.IsVisible)
            {
                m_DeleteTypeOfTextNotesDontUseWpfWindow.Close();
            }

            if (m_AutoCreatePileFromCadWpfWindow != null && m_AutoCreatePileFromCadWpfWindow.IsVisible)
            {
                m_AutoCreatePileFromCadWpfWindow.Close();
            }

            if (m_PlaceElementsFromBlocksCadWpfWindow != null && m_PlaceElementsFromBlocksCadWpfWindow.IsVisible)
            {
                m_PlaceElementsFromBlocksCadWpfWindow.Close();
            }

            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            m_CreateSpaceWpfWindow = null;
            thisApp = this;
            return Result.Succeeded;
        }
        public void ShowCreateSpaceWindow(UIApplication uiapp, CreateSpaceViewModel viewModel)
        {
            if (m_CreateSpaceWpfWindow == null || !m_CreateSpaceWpfWindow.IsVisible)
            {
                CreateSpaceRequestHandler handler = new CreateSpaceRequestHandler(viewModel);
                ExternalEvent exEvent = ExternalEvent.Create(handler);
                m_CreateSpaceWpfWindow = new CreateSpaceWpfWindow(exEvent, handler, viewModel);
                m_CreateSpaceWpfWindow.Show();
            }
        }
        
        public void ShowCopySetFilterFromViewTempViewModel(UIApplication uiapp,
            CopySetOfFilterFromViewTempViewModel viewModel)
        {
            if (m_CopySetOfFilterFromViewTempWpfWindow == null || !m_CopySetOfFilterFromViewTempWpfWindow.IsVisible)
            {
                CopySetOfFilterFromViewTempRequestHandler handler = new CopySetOfFilterFromViewTempRequestHandler(viewModel);
                ExternalEvent exEvent = ExternalEvent.Create(handler);
                m_CopySetOfFilterFromViewTempWpfWindow = new CopySetOfFilterFromViewTempWpfWindow(exEvent, handler, viewModel);
                m_CopySetOfFilterFromViewTempWpfWindow.Show();
            }
        }
        public void ShowDeleteTypeOfTextNotesDontUseViewModel(UIApplication uiapp,
            DeleteTypeOfTextNotesDontUseViewModel viewModel)
        {
            if (m_DeleteTypeOfTextNotesDontUseWpfWindow == null || !m_DeleteTypeOfTextNotesDontUseWpfWindow.IsVisible)
            {
                DeleteTypeOfTextNotesDontUseRequestHandler handler = new DeleteTypeOfTextNotesDontUseRequestHandler(viewModel);
                ExternalEvent exEvent = ExternalEvent.Create(handler);
                m_DeleteTypeOfTextNotesDontUseWpfWindow = new DeleteTypeOfTextNotesDontUseWpfWindow(exEvent, handler, viewModel);
                m_DeleteTypeOfTextNotesDontUseWpfWindow.Show();
            }
        }
        public void ShowSelectElementsVer1ViewModel(UIApplication uiapp, SelectElementsVer1ViewModel viewModel)
        {
            if (m_SelectElementsVer1WpfWindow == null || !m_SelectElementsVer1WpfWindow.IsVisible)
            {
                SelectElementsVer1RequestHandler handler = new SelectElementsVer1RequestHandler(viewModel);
                ExternalEvent exEvent = ExternalEvent.Create(handler);
                m_SelectElementsVer1WpfWindow = new SelectElementsVer1WpfWindow(exEvent, handler, viewModel);
                m_SelectElementsVer1WpfWindow.Show();
            }
        }
        public void ShowSelectElementsViewModel(UIApplication uiapp, SelectElementsViewModel viewModel)
        {
            if (m_SelectElementsWpfWindow == null || !m_SelectElementsWpfWindow.IsVisible)
            {
                SelectElementsRequestHandler handler = new SelectElementsRequestHandler(viewModel);
                ExternalEvent exEvent = ExternalEvent.Create(handler);
                m_SelectElementsWpfWindow = new SelectElementsWpfWindow(exEvent, handler, viewModel);
                m_SelectElementsWpfWindow.Show();
            }
        }
        public void ShowExportSchedulesToExcelViewModel(UIApplication uiApp, ExportSchedulesToExcelViewModel viewModel)
        {
            if (m_ExportSchedulesToExcelWpfWindow == null || !m_ExportSchedulesToExcelWpfWindow.IsVisible)
            {
                ExportSchedulesToExcelRequest request = new ExportSchedulesToExcelRequest();
                ExportSchedulesToExcelRequestHandler handler
                    = new ExportSchedulesToExcelRequestHandler(viewModel, request);
                ExternalEvent exEvent = ExternalEvent.Create(handler);
                m_ExportSchedulesToExcelWpfWindow = new ExportSchedulesToExcelWpfWindow(exEvent, handler, viewModel);
                m_ExportSchedulesToExcelWpfWindow.Show();
            }
        }

        //public void ShowExportLayersWindow(UIApplication uiapp)
        //{
        //    if (m_ExportLayersWindow == null || !m_ExportLayersWindow.IsVisible)
        //    {
        //        var vm = new DWGExportViewModel(uiapp);
        //        var request = new DWGExportRequest();
        //        var handler = new DWGExportRequestHandler(vm, request);
        //        var exEvent = ExternalEvent.Create(handler);
        //        m_ExportLayersWindow = new DWGExportWpfWindow(exEvent, handler, vm);
        //        m_ExportLayersWindow.Show();
        //    }
        //    else
        //    {
        //        m_ExportLayersWindow.Activate();
        //    }
        //}

        public void ShowFindDWGNotUseAndDelViewModel(UIApplication uiapp, FindDWGNotUseAndDelViewModel viewModel)
        {
            if (m_FindDWGNotUseAndDelWpfWindow == null || !m_FindDWGNotUseAndDelWpfWindow.IsVisible)
            {
                FindDWGNotUseAndDelRequestHandler handler = new FindDWGNotUseAndDelRequestHandler(viewModel);
                ExternalEvent exEvent = ExternalEvent.Create(handler);

                m_FindDWGNotUseAndDelWpfWindow = new FindDWGNotUseAndDelWpfWindow(exEvent, handler, viewModel);
                m_FindDWGNotUseAndDelWpfWindow.Show();
            }
        }

        public void ShowAutoCreatePileFromCadViewModel(UIApplication uiapp, AutoCreatePileFromCadViewModel viewModel)
        {
            if (m_AutoCreatePileFromCadWpfWindow == null || !m_AutoCreatePileFromCadWpfWindow.IsVisible)
            {
                AutoCreatePileFromCadRequestHandler handler = new AutoCreatePileFromCadRequestHandler(viewModel);
                ExternalEvent exEvent = ExternalEvent.Create(handler);
                m_AutoCreatePileFromCadWpfWindow = new AutoCreatePileFromCadWpfWindow(exEvent, handler, viewModel);
                m_AutoCreatePileFromCadWpfWindow.Show();
            }
        }

        public void ShowPlaceElementsFromBlocksCadViewModel(UIApplication uiapp,
            PlaceElementsFromBlocksCadViewModel viewModel)
        {
            if (m_PlaceElementsFromBlocksCadWpfWindow == null || !m_PlaceElementsFromBlocksCadWpfWindow.IsVisible)
            {
                PlaceElementsFromBlocksCadRequestHandler handler = new PlaceElementsFromBlocksCadRequestHandler(viewModel);
                ExternalEvent exEvent = ExternalEvent.Create(handler);
                m_PlaceElementsFromBlocksCadWpfWindow = new PlaceElementsFromBlocksCadWpfWindow(exEvent, handler, viewModel);
                m_PlaceElementsFromBlocksCadWpfWindow.Show();
            }
        }

        public void ShowPermissibleRangeFrameViewModel(UIApplication uiapp, PermissibleRangeFrameViewModel viewModel)
        {
            if (m_PermissibleRangeFrameWpfWindow == null || !m_PermissibleRangeFrameWpfWindow.IsVisible)
            {
                PermissibleRangeFrameRequestHandler handler = new PermissibleRangeFrameRequestHandler(viewModel);
                ExternalEvent exEvent = ExternalEvent.Create(handler);
                m_PermissibleRangeFrameWpfWindow = new PermissibleRangeFrameWpfWindow(exEvent, handler, viewModel);
                m_PermissibleRangeFrameWpfWindow.Show();
            }
        }
    }

}
