using System.Collections.Generic;

namespace LanguageCommons.Interfaces
{
    public interface ISettings
    {
        #region Public Properties

        List<string> LastLanguages { get; set; }

        #endregion Public Properties
    }
}