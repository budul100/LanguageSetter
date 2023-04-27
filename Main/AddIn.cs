using Config.Net;
using LanguageCommons.Interfaces;
using NetOffice.OfficeApi;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Tools;
using NetOffice.Tools;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.DryIoc;
using PrismTaskPanes.DryIoc.PowerPoint;
using PrismTaskPanes.Enums;
using PrismTaskPanes.EventArgs;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace LanguageSetter
{
    [COMAddin(AddInName, "Set spell checker language", LoadBehavior.LoadAtStartup),
        ProgId(AddInName),
        Guid("79d02d3d-162b-492e-a6c1-32dc32d999a2"),
        RegistryLocation(RegistrySaveLocation.LocalMachine),
        Codebase,
        ComVisible(true)]
    [CustomUI("RibbonUI.xml", true)]
    [PrismTaskPane(
        id: TaskPaneId,
        title: "Set Language",
        view: typeof(LanguageModule.Views.LanguageView),
        regionName: LanguageRegion,
        width: 300,
        visible: false,
        invisibleAtStart: true,
        scrollBarVertical: ScrollVisibility.Disabled,
        scrollBarHorizontal: ScrollVisibility.Disabled)]
    public class AddIn
        : COMAddin, ITaskPanesReceiver, IDisposable
    {
        #region Private Fields

        private const string AddInName = "LanguageSetter";
        private const string LanguageRegion = "LanguageRegion";
        private const string SettingsFileName = "Settings.json";
        private const string TaskPaneId = "LanguageSetter";

        private bool isDisposed;
        private TaskPanesProvider provider;

        #endregion Private Fields

        #region Public Constructors

        public AddIn()
        {
            OnStartupComplete += OnAddInStart;
        }

        #endregion Public Constructors

        #region Public Methods

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            RegisterFunction(type);
            type.RegisterTaskPaneHost();
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            UnregisterFunction(type);
            type.UnregisterTaskPaneHost();
        }

        public void ClickLanguageSetterToggle(IRibbonControl control, bool isPressed)
        {
            provider.SetTaskPaneVisibility(
                id: TaskPaneId,
                isVisible: isPressed);
        }

        public void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            moduleCatalog.AddModule<LanguageModule.Module>(nameof(LanguageModule));
        }

        public override void CTPFactoryAvailable(object CTPFactoryInst)
        {
            base.CTPFactoryAvailable(CTPFactoryInst);

            provider = new TaskPanesProvider(
                receiver: this,
                officeApplication: Application,
                ctpFactoryInst: CTPFactoryInst,
                showErrorIfAlreadyLoaded: true);

            provider.OnConfigureModuleCatalogEvent += OnConfigureModuleCatalog;
            provider.OnRegisterTypesEvent += OnRegisterTypes;
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        public bool GetLanguageSetterAvailable(IRibbonControl control)
        {
            var result = Application != default
                && Application.Presentations.Count > 0
                && Application.ActivePresentation != default;

            return result;
        }

        public bool GetLanguageSetterPressed(IRibbonControl control)
        {
            return provider.TaskPaneIsVisible(TaskPaneId);
        }

        public void InvalidateRibbonUI()
        {
            RibbonUI?.Invalidate();
        }

        public void OnLoadRibbonUI(NetOffice.OfficeApi.Native.IRibbonUI ribbonUI)
        {
            base.CustomUI_OnLoad(ribbonUI);
        }

        [RegisterErrorHandler]
        public void RegisterErrorHandler(RegisterErrorMethodKind methodKind, Exception exception)
        {
            //Log.Logger.Error(
            //    messageTemplate: exception.Message,
            //    exception: exception);

            Debugger.Break();
        }

        #endregion Public Methods

        #region Protected Methods

        protected virtual void Dispose(bool disposing)
        {
            if (!isDisposed)
            {
                if (disposing)
                {
                    Application.DisposeChildInstances();
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                isDisposed = true;
            }
        }

        #endregion Protected Methods

        #region Private Methods

        private static string GetSettingsPath()
        {
            var result = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                Assembly.GetExecutingAssembly().GetName().Name,
                SettingsFileName);

            return result;
        }

        private void OnAddInStart(ref Array custom)
        {
            Application.PresentationCloseEvent += OnPresentationClose;
        }

        private void OnConfigureModuleCatalog(object sender, ProviderEventArgs<IModuleCatalog> e)
        {
            e.Content.AddModule<LanguageModule.Module>();
        }

        private void OnLanguageSelected(object sender, EventArgs e)
        {
            provider.SetTaskPaneVisibility(
                id: TaskPaneId,
                isVisible: false);
        }

        private void OnPresentationClose(Presentation pres)
        {
            provider.SetTaskPaneVisibility(
                id: TaskPaneId,
                isVisible: false);
        }

        private void OnRegisterTypes(object sender, ProviderEventArgs<IContainerRegistry> e)
        {
            var setter = new SetterService.Service(Application);
            setter.OnSelectedUpdateEvent += OnLanguageSelected;

            e.Content.RegisterInstance<ILanguageSetter>(setter);

            var settingsPath = GetSettingsPath();
            var settings = new ConfigurationBuilder<ILanguageSettings>()
                .UseJsonFile(settingsPath).Build();

            e.Content.RegisterInstance(settings);
        }

        #endregion Private Methods
    }
}