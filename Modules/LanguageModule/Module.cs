using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

namespace LanguageModule
{
    public class Module
        : IModule
    {
        #region Private Fields

        private const string RegionName = "LanguageRegion";

        #endregion Private Fields

        #region Public Methods

        public void OnInitialized(IContainerProvider containerProvider)
        {
            containerProvider.Resolve<IRegionManager>().RegisterViewWithRegion(
                regionName: RegionName,
                viewType: typeof(Views.LanguageView));
        }

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.RegisterForNavigation<Views.LanguageView>();
        }

        #endregion Public Methods
    }
}