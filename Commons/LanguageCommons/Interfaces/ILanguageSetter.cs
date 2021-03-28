using System;
using System.Collections.Generic;

namespace LanguageCommons.Interfaces
{
    public interface ILanguageSetter
    {
        #region Public Events

        event EventHandler OnGivenUpdateEvent;

        event EventHandler OnSelectedUpdateEvent;

        #endregion Public Events

        #region Public Methods

        IEnumerable<int> GetAppLangIds();

        int GetCurrentLangId();

        void SetPresentationLanguage(int languageId);

        void SetSlidesLanguage(int languageId);

        #endregion Public Methods
    }
}