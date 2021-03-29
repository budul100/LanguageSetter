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
                            result = (int)shape.TextFrame.TextRange.LanguageID;
                            break;
                        }
                    }
                }
            }

            return result;
        }

        public void SetPresentationLanguage(int languageId)
        {
            var activePresentation = application.ActivePresentation;

            activePresentation.DefaultLanguageID = (MsoLanguageID)languageId;

            var slides = activePresentation.Slides;

            foreach (var slide in slides)
            {
                SetSlideLanguage(
                    slide: slide,
                    languageId: languageId);
            }

            SetMasterLanguage(
                master: activePresentation.SlideMaster,
                languageId: languageId);

            if (activePresentation.HasTitleMaster == MsoTriState.msoTrue)
            {
                SetMasterLanguage(
                    master: activePresentation.TitleMaster,
                    languageId: languageId);
            }

            if (activePresentation.HasNotesMaster)
            {
                SetMasterLanguage(
                    master: activePresentation.NotesMaster,
                    languageId: languageId);
            }

            if (activePresentation.HasHandoutMaster)
            {
                SetMasterLanguage(
                    master: activePresentation.HandoutMaster,
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

                for (var row = 1; row < table.Rows.Count; row++)
                {
                    for (var column = 1; column < table.Columns.Count; column++)
                    {
                        var cell = table.Cell(
                            row: row,
                            column: column);

                        cell.Shape.TextFrame.TextRange.LanguageID = (MsoLanguageID)languageId;
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

        private void SetMasterLanguage(object master, int languageId)
        {
            var relevant = master as Master;

            if (relevant != default)
            {
                foreach (var relevantShape in relevant.Shapes)
                {
                    SetShapeLanguage(
                        shape: relevantShape,
                        languageId: languageId);
                }

                try
                {
                    foreach (var customLayout in relevant.CustomLayouts)
                    {
                        var relevantLayout = customLayout as CustomLayout;

                        foreach (var relevantShape in relevant.Shapes)
                        {
                            SetShapeLanguage(
                                shape: relevantShape,
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
                foreach (var relevantShape in relevant.Shapes)
                {
                    SetShapeLanguage(
                        shape: relevantShape,
                        languageId: languageId);
                }
            }
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Do not catch general exception types