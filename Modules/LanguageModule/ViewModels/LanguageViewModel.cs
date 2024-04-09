using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using LanguageCommons.Interfaces;
using LanguageCommons.Models;
using LanguageModule.Factories;
using Prism.Commands;
using Prism.Regions;

namespace LanguageModule.ViewModels
{
    public class LanguageViewModel
        : ScopedContentBase
    {
        #region Private Fields

        private const int ListSizeMax = 5;

        private readonly LanguageFactory languageFactory = new();
        private readonly ILanguageSetter languageSetter;
        private readonly ILanguageSettings settings;

        private Language activeLanguage;
        private bool allActivated;
        private Language selectedLanguage;

        #endregion Private Fields

        #region Public Constructors

        public LanguageViewModel(ILanguageSetter languageSetter, ILanguageSettings settings)
        {
            this.languageSetter = languageSetter;
            this.settings = settings;

            languageSetter.OnGivenUpdateEvent += OnLanguageUpdate;

            ActivateLastsCommand = new DelegateCommand(ActivateLasts);
            ActivateAllCommand = new DelegateCommand(ActivateAll);

            SetSlidesLanguageCommand = new DelegateCommand(SetSlidesLanguage);
            SetPresentationLanguageCommand = new DelegateCommand(SetPresentationLanguage);

            settings.LastsSize = GetListSize();
            LastLanguagesSize = settings.LastsSize;

            UpdateAllLanguages();
            SetSelectedLanguage();
        }

        #endregion Public Constructors

        #region Public Properties

        public DelegateCommand ActivateAllCommand { get; }

        public DelegateCommand ActivateLastsCommand { get; }

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

        public ObservableCollection<Language> AllLanguages { get; } = [];

        public ObservableCollection<Language> LastLanguages { get; } = [];

        public int LastLanguagesSize { get; }

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

        public DelegateCommand SetPresentationLanguageCommand { get; }

        public DelegateCommand SetSlidesLanguageCommand { get; }

        #endregion Public Properties

        #region Protected Methods

        protected override void Initialize(NavigationContext navigationContext)
        { }

        #endregion Protected Methods

        #region Private Methods

        private void ActivateAll()
        {
            allActivated = true;
        }

        private void ActivateLasts()
        {
            allActivated = false;
        }

        private IEnumerable<string> GetLastLanguages(Language currentLanguage)
        {
            yield return currentLanguage.Id.ToString();

            if (settings.LastLanguages?.Length > 0)
            {
                var listInd = 1;
                foreach (var language in settings.LastLanguages)
                {
                    if (listInd >= settings.LastsSize) yield break;

                    if (language != currentLanguage.Id.ToString())
                    {
                        yield return language;
                        listInd++;
                    }
                }
            }
        }

        private int GetListSize()
        {
            var result = settings.LastsSize;

            if (result < 1)
            {
                return 1;
            }
            else if (result > ListSizeMax)
            {
                return ListSizeMax;
            }

            return result;
        }

        private void OnLanguageUpdate(object sender, EventArgs e)
        {
            SetSelectedLanguage();
        }

        private void SetActiveLanguage(bool takeSelectedLanguage)
        {
            var currentLanguage = takeSelectedLanguage
                ? SelectedLanguage
                : ActiveLanguage;

            settings.LastLanguages = GetLastLanguages(currentLanguage).ToArray();

            var languages = languageFactory.Get(
                languageIds: settings.LastLanguages,
                size: settings.LastsSize).ToArray();

            LastLanguages.Clear();
            LastLanguages.AddRange(languages);

            ActiveLanguage = currentLanguage;
            SelectedLanguage = currentLanguage;
        }

        private void SetPresentationLanguage()
        {
            SetActiveLanguage(allActivated);

            if (ActiveLanguage != default)
            {
                languageSetter.SetPresentationLanguage(ActiveLanguage.Id);
            }
        }

        private void SetSelectedLanguage()
        {
            var languageId = languageSetter.GetCurrentLangId();
            SelectedLanguage = languageFactory.Get(languageId);

            SetActiveLanguage(true);
        }

        private void SetSlidesLanguage()
        {
            SetActiveLanguage(allActivated);

            if (ActiveLanguage != default)
            {
                languageSetter.SetSlidesLanguage(ActiveLanguage.Id);
            }
        }

        private void UpdateAllLanguages()
        {
            var languageIds = languageSetter
                .GetAppLangIds().ToArray();

            var languages = languageFactory.Get(languageIds)
                .OrderBy(l => l.Name).ToArray();

            foreach (var language in languages)
            {
                AllLanguages.Add(language);
            }
        }

        #endregion Private Methods
    }
}