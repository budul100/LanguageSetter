using LanguageCommons.Interfaces;
using LanguageCommons.Models;
using Prism.Commands;
using Prism.Regions;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace LanguageView.ViewModels
{
    public class LanguageViewModel
        : ScopedContentBase
    {
        #region Private Fields

        private readonly static IEnumerable<CultureInfo> cultures = CultureInfo.GetCultures(CultureTypes.AllCultures);

        private readonly ILanguageSetter languageSetter;
        private readonly ISettings settings;
        private Language activeLanguage;
        private Language selectedLanguage;

        #endregion Private Fields

        #region Public Constructors

        public LanguageViewModel(ILanguageSetter languageSetter, ISettings settings)
        {
            this.languageSetter = languageSetter;
            this.settings = settings;

            languageSetter.OnLanguageUpdateEvent += OnLanguageUpdate;

            SetActiveLanguageCommand = new DelegateCommand(SetActiveLanguage);
            SetSlidesLanguageCommand = new DelegateCommand(SetSlidesLanguage);
            SetPresentationLanguageCommand = new DelegateCommand(SetPresentationLanguage);

            SetLanguages();
            UpdateLanguage();
        }

        #endregion Public Constructors

        #region Public Properties

        public Language ActiveLanguage
        {
            get { return activeLanguage; }
            set
            {
                if (value != default)
                {
                    SetProperty(ref activeLanguage, value);
                }
            }
        }

        public ObservableCollection<Language> AllLanguages { get; set; }

        public ObservableCollection<Language> LastLanguages { get; set; }

        public Language SelectedLanguage
        {
            get { return selectedLanguage; }
            set
            {
                if (value != default)
                {
                    SetProperty(ref selectedLanguage, value);
                }
            }
        }

        public DelegateCommand SetActiveLanguageCommand { get; }

        public DelegateCommand SetPresentationLanguageCommand { get; }

        public DelegateCommand SetSlidesLanguageCommand { get; }

        #endregion Public Properties

        #region Protected Methods

        protected override void Initialize(NavigationContext navigationContext)
        { }

        #endregion Protected Methods

        #region Private Methods

        private Language GetLanguage(int languageId)
        {
            var result = default(Language);

            var culture = cultures
                .Where(c => c.LCID == languageId).SingleOrDefault();

            if (culture != default)
            {
                result = new Language
                {
                    Name = culture.DisplayName,
                    Id = languageId,
                };
            }

            return result;
        }

        private IEnumerable<Language> GetLanguages()
        {
            var langIds = languageSetter
                .GetAppLangIds().ToArray();

            foreach (var languageId in langIds)
            {
                var result = GetLanguage(languageId);

                if (result != default)
                {
                    yield return result;
                }
            }
        }

        private void OnLanguageUpdate(object sender, System.EventArgs e)
        {
            UpdateLanguage();
        }

        private void SetActiveLanguage()
        {
            ActiveLanguage = SelectedLanguage;
        }

        private void SetLanguages()
        {
            var languages = GetLanguages()
                .OrderBy(l => l.Name).ToArray();

            AllLanguages = new ObservableCollection<Language>(languages);
        }

        private void SetPresentationLanguage()
        {
            SetActiveLanguage();

            if (ActiveLanguage != default)
            {
                languageSetter.SetPresentationLanguage(ActiveLanguage.Id);
            }
        }

        private void SetSlidesLanguage()
        {
            SetActiveLanguage();

            if (ActiveLanguage != default)
            {
                languageSetter.SetSlidesLanguage(ActiveLanguage.Id);
            }
        }

        private void UpdateLanguage()
        {
            var languageId = languageSetter.GetCurrentLangId();

            SelectedLanguage = GetLanguage(languageId);

            SetActiveLanguage();
        }

        #endregion Private Methods
    }
}