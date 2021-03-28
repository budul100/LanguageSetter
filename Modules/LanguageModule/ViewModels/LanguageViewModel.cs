using LanguageCommons.Interfaces;
using LanguageCommons.Models;
using LanguageModule.Factories;
using Prism.Commands;
using Prism.Regions;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace LanguageModule.ViewModels
{
    public class LanguageViewModel
        : ScopedContentBase
    {
        #region Private Fields

        private const int ListSizeMax = 5;

        private readonly LanguageFactory languageFactory = new LanguageFactory();
        private readonly ILanguageSetter languageSetter;
        private readonly ILanguageSettings settings;

        private Language activeLanguage;
        private Language selectedLanguage;

        #endregion Private Fields

        #region Public Constructors

        public LanguageViewModel(ILanguageSetter languageSetter, ILanguageSettings settings)
        {
            this.languageSetter = languageSetter;
            this.settings = settings;

            languageSetter.OnLanguageUpdateEvent += OnLanguageUpdate;

            SetSlidesLanguageCommand = new DelegateCommand(SetSlidesLanguage);
            SetPresentationLanguageCommand = new DelegateCommand(SetPresentationLanguage);

            settings.LastsSize = GetListSize();
            LastLanguagesSize = settings.LastsSize;

            UpdateAllLanguages();
            SetSelectedLanguage();
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

        public ObservableCollection<Language> AllLanguages { get; } = new ObservableCollection<Language>();

        public ObservableCollection<Language> LastLanguages { get; } = new ObservableCollection<Language>();

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

        private IEnumerable<string> GetLastLanguages()
        {
            yield return SelectedLanguage.Id.ToString();

            if (settings.LastLanguages?.Any() ?? false)
            {
                var listInd = 1;
                foreach (var language in settings.LastLanguages)
                {
                    if (listInd > settings.LastsSize) yield break;

                    if (language != SelectedLanguage.Id.ToString())
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
                result = 1;
            }
            else if (result > ListSizeMax)
            {
                result = ListSizeMax;
            }

            return result;
        }

        private void OnLanguageUpdate(object sender, EventArgs e)
        {
            SetSelectedLanguage();
        }

        private void SetActiveLanguage()
        {
            settings.LastLanguages = GetLastLanguages().ToArray();

            var languages = languageFactory.Get(
                languageIds: settings.LastLanguages,
                size: settings.LastsSize).ToArray();

            LastLanguages.Clear();

            foreach (var language in languages)
            {
                LastLanguages.Add(language);
            }

            ActiveLanguage = SelectedLanguage;
        }

        private void SetPresentationLanguage()
        {
            SetActiveLanguage();

            if (ActiveLanguage != default)
            {
                languageSetter.SetPresentationLanguage(ActiveLanguage.Id);
            }
        }

        private void SetSelectedLanguage()
        {
            var languageId = languageSetter.GetCurrentLangId();
            SelectedLanguage = languageFactory.Get(languageId);

            SetActiveLanguage();
        }

        private void SetSlidesLanguage()
        {
            SetActiveLanguage();

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