using LanguageCommons.Interfaces;
using LanguageSetterApp.Views;
using Prism.Ioc;
using Prism.Modularity;
using System.Collections.Generic;
using System.Windows;

namespace LanguageSetterApp
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App
        : ILanguageSetter
    {
        #region Public Methods

        public IEnumerable<int> GetAppLangIds()
        {
            throw new System.NotImplementedException();
        }

        public int GetCurrentLangId()
        {
            throw new System.NotImplementedException();
        }

        public static IEnumerable<int> GetLanguageIds()
        {
            for (int i = 1000; i < 3000; i++)
            {
                yield return i;
            }
        }

        public void SetPresentationLanguage(int languageId)
        {
        }

        public void SetSlidesLanguage(int languageId)
        {
        }

        #endregion Public Methods

        #region Protected Methods

        protected override void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            moduleCatalog.AddModule<LanguageView.Module>(nameof(LanguageView));
        }

        protected override Window CreateShell()
        {
            return Container.Resolve<MainWindow>();
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.RegisterInstance<ILanguageSetter>(this);
        }

        #endregion Protected Methods
    }
}