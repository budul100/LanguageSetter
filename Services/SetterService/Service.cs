using LanguageCommons.Interfaces;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi;
using SetterService.Extensions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace SetterService
{
    public class Service
        : ILanguageSetter
    {
        #region Private Fields

        private const string SlideRangeName = "PowerPoint.SlideRange";

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

            var childObjects = selection.ChildObjects;
            var hasSlides = childObjects.Any(c => c.InstanceFriendlyName == SlideRangeName);

            if (hasSlides)
            {
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
                slide.SetSlideLanguage(languageId);
            }

            relevant.SlideMaster.SetMasterLanguage(languageId);

            if (relevant.HasTitleMaster == MsoTriState.msoTrue)
            {
                relevant.TitleMaster.SetMasterLanguage(languageId);
            }

            if (relevant.HasNotesMaster)
            {
                relevant.NotesMaster.SetMasterLanguage(languageId);
            }

            if (relevant.HasHandoutMaster)
            {
                relevant.HandoutMaster.SetMasterLanguage(languageId);
            }

            var designs = relevant.Designs;

            foreach (var design in designs)
            {
                design.SetDesignLanguage(languageId);
            }

            OnSelectedUpdateEvent?.Invoke(
                sender: this,
                e: default);
        }

        public void SetSlidesLanguage(int languageId)
        {
            var activeWindow = application.ActiveWindow;
            var selection = activeWindow.Selection;

            var childObjects = selection.ChildObjects;
            var hasSlides = childObjects.Any(c => c.InstanceFriendlyName == SlideRangeName);

            if (hasSlides)
            {
                var slides = selection.SlideRange;

                foreach (var slide in slides)
                {
                    slide.SetSlideLanguage(languageId);
                }

                OnSelectedUpdateEvent?.Invoke(
                    sender: this,
                    e: default);
            }
        }

        #endregion Public Methods

        #region Private Methods

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

        #endregion Private Methods
    }
}