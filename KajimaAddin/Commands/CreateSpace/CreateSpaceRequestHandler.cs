using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Mechanical;
using UnitUtils = SKToolsAddins.Utils.UnitUtils;
using System.Windows.Controls;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB.Structure;
using SKToolsAddins.ViewModel;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Document = Autodesk.Revit.DB.Document;

namespace SKToolsAddins.Commands.CreateSpace
{
    public class CreateSpaceRequestHandler : IExternalEventHandler
    {
        public Document Doc;
        private CreateSpaceViewModel ViewModel;

        public CreateSpaceRequestHandler(CreateSpaceViewModel viewModel)
        {
            ViewModel = viewModel;
        }

        private CreateSpaceRequest m_Request = new CreateSpaceRequest();
        public CreateSpaceRequest Request
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
                    case (RequestId.CreateSpace):
                        CreateSpaces(uiapp, ViewModel);
                        break;
                    case RequestId.DeleteSpace:
                        DeleteSpaces(uiapp, ViewModel);
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
            return "スペースー括作成";
        }

        #region CreateSpaces
        public void CreateSpaces(UIApplication uiapp, CreateSpaceViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var tagPlacementBOX = viewModel.TagPlacementBOX;
            var nameNumberBOX = viewModel.NameNumberBOX;
            var spaceOffsetBOX = viewModel.SpaceOffsetBOX;
            var setSpaceOffet = viewModel.SetSpaceOffet;
            var selTagTypeSpace = viewModel.SelTagTypeSpace;
            var selectedViews = viewModel.SelectedViews;
            var selPhase = viewModel.SelPhase;
            ICollection<ElementId> spaces;

            foreach (var selView in selectedViews)
            {
                Level level = selView.GenLevel; // Giả sử view là một plan view
                if (level != null)
                {
                    using (Transaction tx = new Transaction(doc, "Create Space"))
                    {
                        tx.Start();
                        var failureOptions = tx.GetFailureHandlingOptions();
                        failureOptions.SetFailuresPreprocessor(new DuplicateNumberDisable());
                        tx.SetFailureHandlingOptions(failureOptions);
                        spaces = doc.Create.NewSpaces2(level, selPhase, selView);

                        using (SubTransaction deleteTagTx = new SubTransaction(doc))
                        {
                            var listTagSpace = new ObservableCollection<SpaceTag>(new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_MEPSpaceTags)
                                .WhereElementIsNotElementType()
                                .Cast<SpaceTag>()
                                .ToList());
                            deleteTagTx.Start();
                            if (listTagSpace != null)
                            {
                                List<ElementId> tagIdsToDelete = new List<ElementId>();
                                foreach (Element element in listTagSpace)
                                {
                                    tagIdsToDelete.Add(element.Id);
                                }
                                foreach (ElementId tagId in tagIdsToDelete)
                                {
                                    doc.Delete(tagId);
                                }
                            }
                            deleteTagTx.Commit();
                        }
                        if (tagPlacementBOX)
                        {
                            using (SubTransaction createTagtx = new SubTransaction(doc))
                            {
                                createTagtx.Start();
                                foreach (var exisSpace in spaces)
                                {
                                    Space space = doc.GetElement(exisSpace) as Space;
                                    XYZ tagPosition = GetSpaceLocationPoint(space);
                                    Reference refTag = new Reference(space);
                                    var tagID = selTagTypeSpace.Id;
                                    IndependentTag newTag = IndependentTag.Create(doc, tagID, selView.Id, refTag, false, TagOrientation.Horizontal, tagPosition);
                                }
                                createTagtx.Commit();
                            }
                        }
                        if (spaceOffsetBOX)
                        {
                            using (SubTransaction changeLimittx = new SubTransaction(doc))
                            {
                                changeLimittx.Start();
                                foreach (var exisSpace in spaces)
                                {
                                    Space space = doc.GetElement(exisSpace) as Space;
                                    space.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(Utils.UnitUtils.MmToFeet(setSpaceOffet));
                                }
                                changeLimittx.Commit();
                            }
                        }
                        if (nameNumberBOX)
                        {
                            using (SubTransaction copyParatx = new SubTransaction(doc))
                            {
                                copyParatx.Start();
                                foreach (var exisSpace in spaces)
                                {
                                    Space space = doc.GetElement(exisSpace) as Space;
                                    var nameRoom = space.get_Parameter(BuiltInParameter.SPACE_ASSOC_ROOM_NAME).AsValueString();
                                    space.get_Parameter(BuiltInParameter.ROOM_NAME).Set(string.Empty);
                                    space.get_Parameter(BuiltInParameter.ROOM_NAME).Set(nameRoom);
                                    var numRoom = space.get_Parameter(BuiltInParameter.SPACE_ASSOC_ROOM_NUMBER).AsValueString();
                                    space.get_Parameter(BuiltInParameter.ROOM_NUMBER).Set(numRoom);
                                }
                                copyParatx.Commit();
                            }
                        }
                        tx.Commit();
                    }
                }
                else
                {
                    TaskDialog.Show("Create Space", "Target level not found.");
                }
            }
        }
        #endregion

        #region DeleteSpaces
        public void DeleteSpaces(UIApplication uiapp, CreateSpaceViewModel viewModel)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            // Get list exis Space
            var ListExisSpaces = viewModel.ListExistSpaces;
            if (ListExisSpaces != null)
            {
                List<ElementId> spaceIdsToDelete = new List<ElementId>();
                foreach (Element element in ListExisSpaces)
                {
                    if (element is Space space)
                    {
                        spaceIdsToDelete.Add(space.Id);
                    }
                }
                using (Transaction deleteSpacestx = new Transaction(doc, "Delete Spaces"))
                {
                    deleteSpacestx.Start();
                    foreach (ElementId spaceId in spaceIdsToDelete)
                    {
                        doc.Delete(spaceId);
                    }
                    deleteSpacestx.Commit();
                }
            }
        }
        #endregion

        private XYZ GetSpaceLocationPoint(Space space)
        {
            // Check if the space has a location point
            if (space.Location is LocationPoint locationPoint)
            {
                return locationPoint.Point;
            }
            return null;
        }
    }

    public class DuplicateNumberDisable : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
        {
            var failures = a.GetFailureMessages();
            foreach (var f in failures)
            {
                var id = f.GetFailureDefinitionId();
                if (BuiltInFailures.GeneralFailures.DuplicateValue == id)
                {
                    a.DeleteWarning(f);
                }
            }
            return FailureProcessingResult.Continue;
        }
    }
}
