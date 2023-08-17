using NetOffice.Exceptions;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi;

namespace SetterService.Extensions
{
    internal static class ObjectExtensions
    {
        #region Public Methods

        public static void SetDesignLanguage(this object design, int languageId)
        {
            var relevant = design as Design;

            if (relevant != default)
            {
                relevant.SlideMaster.SetMasterLanguage(languageId);

                if (relevant.HasTitleMaster == MsoTriState.msoTrue)
                {
                    relevant.TitleMaster.SetMasterLanguage(languageId);
                }
            }
        }

        public static void SetMasterLanguage(this object master, int languageId)
        {
            var relevant = master as Master;

            if (relevant != default)
            {
                foreach (var shape in relevant.Shapes)
                {
                    shape.SetShapeLanguage(languageId);
                }

                try
                {
                    foreach (var customLayout in relevant.CustomLayouts)
                    {
                        var relevantLayout = customLayout as CustomLayout;

                        foreach (var shape in relevantLayout.Shapes)
                        {
                            shape.SetShapeLanguage(languageId);
                        }
                    }
                }
                catch (PropertyGetCOMException)
                { }
            }
        }

        public static void SetSlideLanguage(this object slide, int languageId)
        {
            var relevant = slide as Slide;

            if (relevant != default)
            {
                foreach (var shape in relevant.Shapes)
                {
                    shape.SetShapeLanguage(languageId);
                }

                if (relevant.HasNotesPage == MsoTriState.msoTrue)
                {
                    var notesPage = relevant.NotesPage;

                    foreach (var shape in notesPage.Shapes)
                    {
                        shape.SetShapeLanguage(languageId);
                    }
                }
            }
        }

        #endregion Public Methods

        #region Private Methods

        private static void SetShapeLanguage(this Shape shape, int languageId)
        {
            if (shape.HasTextFrame == MsoTriState.msoTrue)
            {
                var textFrame = shape.TextFrame;

                if (textFrame.HasText == MsoTriState.msoTrue)
                {
                    try
                    {
                        textFrame.TextRange.LanguageID = (MsoLanguageID)languageId;
                    }
                    catch
                    { }
                }
            }
            else if (shape.HasTable == MsoTriState.msoTrue)
            {
                var table = shape.Table;

                var rows = table.Rows;

                for (var rowIndex = 1; rowIndex < rows.Count; rowIndex++)
                {
                    var columns = table.Columns;

                    for (var columnIndex = 1; columnIndex < columns.Count; columnIndex++)
                    {
                        var cell = table.Cell(
                            row: rowIndex,
                            column: columnIndex);

                        var cellShape = cell.Shape;

                        SetShapeLanguage(
                            shape: cellShape,
                            languageId: languageId);
                    }
                }
            }
            else if (shape.Type == MsoShapeType.msoGroup || shape.Type == MsoShapeType.msoSmartArt)
            {
                foreach (var groupItem in shape.GroupItems)
                {
                    SetShapeLanguage(
                        shape: groupItem,
                        languageId: languageId);
                }
            }
        }

        #endregion Private Methods
    }
}