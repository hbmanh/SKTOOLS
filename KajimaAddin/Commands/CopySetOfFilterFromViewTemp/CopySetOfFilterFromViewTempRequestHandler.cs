using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SKToolsAddins.ViewModel;
using Document = Autodesk.Revit.DB.Document;

namespace SKToolsAddins.Commands.CopySetOfFilterFromViewTemp
{
    public class CopySetOfFilterFromViewTempRequestHandler : IExternalEventHandler
    {
        public Document Doc;
        private CopySetOfFilterFromViewTempViewModel ViewModel;

        public CopySetOfFilterFromViewTempRequestHandler(CopySetOfFilterFromViewTempViewModel viewModel)
        {
            ViewModel = viewModel;
        }

        private CopySetOfFilterFromViewTempRequest m_Request = new CopySetOfFilterFromViewTempRequest();

        public CopySetOfFilterFromViewTempRequest Request
        {
            get { return m_Request; }
        }

        public void Execute(UIApplication uiapp)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        break;
                    case (RequestId.OK):
                        CopyFilterOptionData(uiapp, ViewModel);
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("エラー", ex.Message);
            }
        }

        public string GetName()
        {
            return "フィルター色コピー";
        }

        #region Copy Filter Option Data
        public void CopyFilterOptionData(UIApplication uiapp, CopySetOfFilterFromViewTempViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var allCopyBOX = viewModel.AllCopyBOX;
            var patternCopyBOX = viewModel.PatternCopyBOX;
            var cutSetCopyBOX = viewModel.CutSetCopyBOX;
            var selViewTarget = viewModel.SelViewTarget;
            var selViewTemplate = viewModel.SelViewTemplate;
            var selFiltersId = viewModel.SelFilter;
            var filterOverrides = new Dictionary<ElementId, OverrideGraphicSettings>();

            bool filterVisile;
            bool filterEnable;

            OverrideGraphicSettings overrideAllSettings;
            OverrideGraphicSettings overridePatternSettings = new OverrideGraphicSettings();
            OverrideGraphicSettings overrideCutSettings = new OverrideGraphicSettings();

            /// Get filter overrides from selected filters
            if (selFiltersId != null)
            {
                foreach (var selFilterId in selFiltersId)
                {
                    filterVisile = selViewTemplate.GetFilterVisibility(selFilterId.FilterId);
                    filterEnable = selViewTemplate.GetIsFilterEnabled(selFilterId.FilterId);
                    /// Get the filter overrides
                    overrideAllSettings = selViewTemplate.GetFilterOverrides(selFilterId.FilterId);
                    ///Setting Override color
                    if (overrideAllSettings != null)
                    {
                        ///Projection|Surface

                        overridePatternSettings.SetProjectionLinePatternId(overrideAllSettings.ProjectionLinePatternId);
                        overridePatternSettings.SetProjectionLineColor(overrideAllSettings.ProjectionLineColor);
                        overridePatternSettings.SetProjectionLineWeight(overrideAllSettings.ProjectionLineWeight);

                        overridePatternSettings.SetSurfaceForegroundPatternVisible(overrideAllSettings.IsSurfaceForegroundPatternVisible);
                        overridePatternSettings.SetSurfaceForegroundPatternId(overrideAllSettings.SurfaceForegroundPatternId);
                        overridePatternSettings.SetSurfaceForegroundPatternColor(overrideAllSettings.SurfaceForegroundPatternColor);
                        overridePatternSettings.SetSurfaceBackgroundPatternVisible(overrideAllSettings.IsSurfaceBackgroundPatternVisible);
                        overridePatternSettings.SetSurfaceBackgroundPatternId(overrideAllSettings.SurfaceBackgroundPatternId);
                        overridePatternSettings.SetSurfaceBackgroundPatternColor(overrideAllSettings.SurfaceBackgroundPatternColor);

                        overridePatternSettings.SetSurfaceTransparency(overrideAllSettings.Transparency);
                        overridePatternSettings.SetHalftone(overrideAllSettings.Halftone);

                        ///Cut

                        overrideCutSettings.SetCutLinePatternId(overrideAllSettings.CutLinePatternId);
                        overrideCutSettings.SetCutLineColor(overrideAllSettings.CutLineColor);
                        overrideCutSettings.SetCutLineWeight(overrideAllSettings.CutLineWeight);

                        overrideCutSettings.SetCutForegroundPatternVisible(overrideAllSettings.IsCutForegroundPatternVisible);
                        overrideCutSettings.SetCutForegroundPatternId(overrideAllSettings.CutForegroundPatternId);
                        overrideCutSettings.SetCutForegroundPatternColor(overrideAllSettings.CutForegroundPatternColor);

                        overrideCutSettings.SetCutBackgroundPatternVisible(overrideAllSettings.IsCutBackgroundPatternVisible);
                        overrideCutSettings.SetCutBackgroundPatternId(overrideAllSettings.CutBackgroundPatternId);
                        overrideCutSettings.SetCutBackgroundPatternColor(overrideAllSettings.CutBackgroundPatternColor);

                        overrideCutSettings.SetHalftone(overrideAllSettings.Halftone);
                    }

                    /// Store the filter and its overrides in the dictionary
                    filterOverrides[selFilterId.FilterId] = overrideAllSettings;

                    if (allCopyBOX)
                    {
                        foreach (var view in selViewTarget)
                        {
                            View templateView = doc.GetElement(view.ViewTemplateId) as View;

                            if (templateView == null)
                            {
                                using (Transaction transaction = new Transaction(doc))
                                {
                                    transaction.Start("Apply All Filter Overrides to View");

                                    try
                                    {
                                        if (!view.GetFilters().Contains(selFilterId.FilterId))
                                        {
                                            view.AddFilter(selFilterId.FilterId);
                                            view.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            view.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            view.SetFilterOverrides(selFilterId.FilterId, overrideAllSettings);
                                        }
                                        else
                                        {
                                            view.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            view.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            view.SetFilterOverrides(selFilterId.FilterId, overrideAllSettings);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("Error", "Failed to add filter: " + ex.Message);
                                    }
                                    transaction.Commit();
                                }
                            }
                            else
                            {
                                using (Transaction transaction = new Transaction(doc))
                                {
                                    transaction.Start("Apply All Filter Overrides to View");
                                    try
                                    {
                                        if (!view.GetFilters().Contains(selFilterId.FilterId))
                                        {
                                            templateView.AddFilter(selFilterId.FilterId);
                                            templateView.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            templateView.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            templateView.SetFilterOverrides(selFilterId.FilterId, overrideAllSettings);
                                        }
                                        else
                                        {
                                            templateView.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            templateView.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            templateView.SetFilterOverrides(selFilterId.FilterId, overrideAllSettings);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("Error", "Failed to add filter: " + ex.Message);
                                    }
                                    transaction.Commit();
                                }
                            }
                        }
                    }

                    if (patternCopyBOX)
                    {
                        foreach (var view in selViewTarget)
                        {
                            View templateView = doc.GetElement(view.ViewTemplateId) as View;

                            if (templateView == null)
                            {
                                using (Transaction transaction = new Transaction(doc))
                                {
                                    transaction.Start("Apply All Filter Overrides to View");

                                    try
                                    {
                                        if (!view.GetFilters().Contains(selFilterId.FilterId))
                                        {
                                            view.AddFilter(selFilterId.FilterId);
                                            view.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            view.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            view.SetFilterOverrides(selFilterId.FilterId, overridePatternSettings);
                                        }
                                        else
                                        {
                                            view.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            view.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            view.SetFilterOverrides(selFilterId.FilterId, overridePatternSettings);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("Error", "Failed to add filter: " + ex.Message);
                                    }
                                    transaction.Commit();
                                }
                            }
                            else
                            {
                                using (Transaction transaction = new Transaction(doc))
                                {
                                    transaction.Start("Apply All Filter Overrides to View");
                                    try
                                    {
                                        if (!view.GetFilters().Contains(selFilterId.FilterId))
                                        {
                                            templateView.AddFilter(selFilterId.FilterId);
                                            templateView.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            templateView.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            templateView.SetFilterOverrides(selFilterId.FilterId, overridePatternSettings);
                                        }
                                        else
                                        {
                                            templateView.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            templateView.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            templateView.SetFilterOverrides(selFilterId.FilterId, overridePatternSettings);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("Error", "Failed to add filter: " + ex.Message);
                                    }
                                    transaction.Commit();
                                }
                            }
                        }
                    }
                    
                    if (cutSetCopyBOX)
                    {
                        foreach (var view in selViewTarget)
                        {
                            View templateView = doc.GetElement(view.ViewTemplateId) as View;

                            if (templateView == null)
                            {
                                using (Transaction transaction = new Transaction(doc))
                                {
                                    transaction.Start("Apply All Filter Overrides to View");

                                    try
                                    {
                                        if (!view.GetFilters().Contains(selFilterId.FilterId))
                                        {
                                            view.AddFilter(selFilterId.FilterId);
                                            view.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            view.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            view.SetFilterOverrides(selFilterId.FilterId, overrideCutSettings);
                                        }
                                        else
                                        {
                                            view.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            view.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            view.SetFilterOverrides(selFilterId.FilterId, overrideCutSettings);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("Error", "Failed to add filter: " + ex.Message);
                                    }
                                    transaction.Commit();
                                }
                            }
                            else
                            {
                                using (Transaction transaction = new Transaction(doc))
                                {
                                    transaction.Start("Apply All Filter Overrides to View");
                                    try
                                    {
                                        if (!view.GetFilters().Contains(selFilterId.FilterId))
                                        {
                                            templateView.AddFilter(selFilterId.FilterId);
                                            templateView.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            templateView.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            templateView.SetFilterOverrides(selFilterId.FilterId, overrideCutSettings);
                                        }
                                        else
                                        {
                                            templateView.SetIsFilterEnabled(selFilterId.FilterId, filterEnable);
                                            templateView.SetFilterVisibility(selFilterId.FilterId, filterVisile);
                                            templateView.SetFilterOverrides(selFilterId.FilterId, overrideCutSettings);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("Error", "Failed to add filter: " + ex.Message);
                                    }
                                    transaction.Commit();
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                TaskDialog.Show("Warning", "No Selected Filters.");
            }
        }
        #endregion
    }
}
