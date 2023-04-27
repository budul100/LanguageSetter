using LanguageCommons.Interfaces;
using LanguageModule.ViewModels;
using Moq;
using System;
using System.Collections.Generic;
using Xunit;

namespace LanguageModuleTests.ViewModels
{
    public class LanguageViewModelFixture
    {
        #region Private Fields

        private readonly Mock<ILanguageSetter> setterMock;
        private readonly Mock<ILanguageSettings> settingsMock;

        private int languageId;

        #endregion Private Fields

        #region Public Constructors

        public LanguageViewModelFixture()
        {
            setterMock = new Mock<ILanguageSetter>();
            setterMock.Setup(x => x.GetAppLangIds()).Returns(GetLanguageIds);
            setterMock.Setup(x => x.GetCurrentLangId()).Returns(GetLanguageId);

            settingsMock = new Mock<ILanguageSettings>();
        }

        #endregion Public Constructors

        #region Public Methods

        [Fact]
        public void ActualINotifyPropertyChangedCalled()
        {
            var vm = new LanguageViewModel(setterMock.Object, settingsMock.Object);

            Assert.PropertyChanged(vm, nameof(vm.ActiveLanguage), () => setterMock.Raise(x => x.OnGivenUpdateEvent += default, default(EventArgs)));
        }

        [Fact]
        public void AllLanguagesLoaded()
        {
            var vm = new LanguageViewModel(setterMock.Object, settingsMock.Object);

            setterMock.Verify(x => x.GetAppLangIds(), Times.Once);

            Assert.True(vm.AllLanguages.Count > 100);
        }

        [Fact]
        public void SelectedINotifyPropertyChangedCalled()
        {
            var vm = new LanguageViewModel(setterMock.Object, settingsMock.Object);

            Assert.PropertyChanged(vm, nameof(vm.SelectedLanguage), () => setterMock.Raise(x => x.OnGivenUpdateEvent += default, default(EventArgs)));
        }

        #endregion Public Methods

        #region Private Methods

        private int GetLanguageId()
        {
            languageId = languageId == 1031
                ? 1032
                : 1031;

            return languageId;
        }

        private IEnumerable<int> GetLanguageIds()
        {
            for (int i = 1000; i < 3000; i++)
            {
                yield return i;
            }
        }

        #endregion Private Methods
    }
}