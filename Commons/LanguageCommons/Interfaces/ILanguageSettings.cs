using Config.Net;

namespace LanguageCommons.Interfaces
{
    public interface ILanguageSettings
    {
        #region Public Properties

        string[] LastLanguages { get; set; }

        [Option(DefaultValue = 3)]
        int LastsSize { get; set; }

        #endregion Public Properties
    }
}