#pragma warning disable CA1031 // Do not catch general exception types

using LanguageCommons.Interfaces;
using NetOffice.Exceptions;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace LanguageService
{
    public class Service
        : ILanguageSetter
    {
        #region Private Fields

        private readonly Application application;

        #endregion Private Fields

        #region Public Constructors

        public Service(Application application)
        {
            this.application = application;

            application.AfterNewPresentationEvent += OnPresentationNew;
            application.AfterPresentationOpenEvent += OnPresentationOpen;
        }

        #endregion Public Constructors

        #region Public Events

        public event EventHandler OnGivenUpdateEvent;

        public event EventHandler OnSelectedUpdateEvent;

        #endregion Public Events

        #region Public Methods

        public IEnumerable<int> GetAppLangIds()
        {
            foreach (var msoLanguageId in Enum.GetValues(typeof(MsoLanguageID)))
            {
                var languageId = (int)msoLanguageId;

                if (languageId > 0)
                {
                    yield return languageId;
                }
            }
        }

        public int GetCurrentLangId()
        {
            var result = CultureInfo.InstalledUICulture.LCID;

            var activeWindow = application.ActiveWindow;
            var selection = activeWindow.Selection;
            var slides = selection.SlideRange;

            foreach (var slide in slides)
            {
                if (slide.Shapes != default)
                {
                    foreach (var shape in slide.Shapes)
                    {
                        if (shape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            var textFrame = shape.TextFrame;
                            var textRange = textFrame.TextRange;
                            result = (int)textRange.LanguageID;

                            break;
                        }
                    }
                }
            }

            return result;
        }

        public void SetPresentationLanguage(int languageId)
        {
            var relevant = application.ActivePresentation;

            relevant.DefaultLanguageID = (MsoLanguageID)languageId;

            var slides = relevant.Slides;

            foreach (var slide in slides)
            {
                SetSlideLanguage(
                    slide: slide,
                    languageId: languageId);
            }

            SetMasterLanguage(
                master: relevant.SlideMaster,
                languageId: languageId);

            if (relevant.HasTitleMaster == MsoTriState.msoTrue)
            {
                SetMasterLanguage(
                    master: relevant.TitleMaster,
                    languageId: languageId);
            }

            if (relevant.HasNotesMaster)
            {
                SetMasterLanguage(
                    master: relevant.NotesMaster,
                    languageId: languageId);
            }

            if (relevant.HasHandoutMaster)
            {
                SetMasterLanguage(
                    master: relevant.HandoutMaster,
                    languageId: languageId);
            }

            var designs = relevant.Designs;

            foreach (var design in designs)
            {
                SetDesignLanguage(
                    design: design,
                    languageId: languageId);
            }

            OnSelectedUpdateEvent?.Invoke(
                sender: this,
                e: default);
        }

        public void SetSlidesLanguage(int languageId)
        {
            var activeWindow = application.ActiveWindow;
            var selection = activeWindow.Selection;
            var slides = selection.SlideRange;

            foreach (var slide in slides)
            {
                SetSlideLanguage(
                    slide: slide,
                    languageId: languageId);
            }

            OnSelectedUpdateEvent?.Invoke(
                sender: this,
                e: default);
        }

        #endregion Public Methods

        #region Private Methods

        private static void SetShapeLanguage(Shape shape, int languageId)
        {
            if (shape.HasTextFrame == MsoTriState.msoTrue)
            {
                shape.TextFrame.TextRange.LanguageID = (MsoLanguageID)languageId;
            }

            if (shape.HasTable == MsoTriState.msoTrue)
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
                        var textFrame = cellShape.TextFrame;
                        var textRange = textFrame.TextRange;

                        textRange.LanguageID = (MsoLanguageID)languageId;
                    }
                }
            }

            if (shape.Type == MsoShapeType.msoGroup || shape.Type == MsoShapeType.msoSmartArt)
            {
                foreach (var groupItem in shape.GroupItems)
                {
                    SetShapeLanguage(
                        shape: groupItem,
                        languageId: languageId);
                }
            }
        }

        private void OnPresentationNew(Presentation pres)
        {
            OnGivenUpdateEvent?.Invoke(
                sender: this,
                e: default);
        }

        private void OnPresentationOpen(Presentation pres)
        {
            OnGivenUpdateEvent?.Invoke(
                sender: this,
                e: default);
        }

        private void SetDesignLanguage(object design, int languageId)
        {
            var relevant = design as Design;

            if (relevant != default)
            {
                SetMasterLanguage(
                    master: relevant.SlideMaster,
                    languageId: languageId);

                if (relevant.HasTitleMaster == MsoTriState.msoTrue)
                {
                    SetMasterLanguage(
                        master: relevant.TitleMaster,
                        languageId: languageId);
                }
            }
        }

        private void SetMasterLanguage(object master, int languageId)
        {
            var relevant = master as Master;

            if (relevant != default)
            {
                foreach (var shape in relevant.Shapes)
                {
                    SetShapeLanguage(
                        shape: shape,
                        languageId: languageId);
                }

                try
                {
                    foreach (var customLayout in relevant.CustomLayouts)
                    {
                        var relevantLayout = customLayout as CustomLayout;

                        foreach (var shape in relevantLayout.Shapes)
                        {
                            SetShapeLanguage(
                                shape: shape,
                                languageId: languageId);
                        }
                    }
                }
                catch (PropertyGetCOMException)
                { }
            }
        }

        private void SetSlideLanguage(object slide, int languageId)
        {
            var relevant = slide as Slide;

            if (relevant != default)
            {
                foreach (var shape in relevant.Shapes)
                {
                    SetShapeLanguage(
                        shape: shape,
                        languageId: languageId);
                }

                if (relevant.HasNotesPage == MsoTriState.msoTrue)
                {
                    var notesPage = relevant.NotesPage;

                    foreach (var shape in notesPage.Shapes)
                    {
                        SetShapeLanguage(
                            shape: shape,
                            languageId: languageId);
                    }
                }
            }
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Do not catch general exception types