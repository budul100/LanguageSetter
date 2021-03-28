#pragma warning disable CA1031 // Do not catch general exception types

using Config.Net;
using LanguageCommons.Interfaces;
using NetOffice.Exceptions;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Tools;
using NetOffice.Tools;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.Commons.Enums;
using PrismTaskPanes.DryIoc;
using PrismTaskPanes.Interfaces;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
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
        ScrollBarVertical = ScrollVisibility.Disabled,
        ScrollBarHorizontal = ScrollVisibility.Disabled)]
    public class AddIn
        : COMAddin, ITaskPanesReceiver, ILanguageSetter, IDisposable
    {
        #region Private Fields

        private const string AddInName = "LanguageSetter";
        private const string LanguageRegion = "LanguageRegion";
        private const string SettingsFileName = "Settings.json";
        private const string TaskPaneId = "LanguageSetter";

        private bool isDisposed;

        #endregion Private Fields

        #region Public Constructors

        public AddIn()
        {
            OnStartupComplete += OnAddInStart;
        }

        #endregion Public Constructors

        #region Public Events

        public event EventHandler OnLanguageUpdateEvent;

        #endregion Public Events

        #region Public Methods

        public void ClickLanguageSetterToggle(IRibbonControl control, bool isPressed)
        {
            this.SetTaskPaneVisible(
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

            this.InitializeProvider(
                application: Application,
                ctpFactoryInst: CTPFactoryInst);
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        public IEnumerable<int> GetAppLangIds()
        {
            foreach (var msoLanguageId in Enum.GetValues(typeof(MsoLanguageID)))
            {
                var languageId = (int)msoLanguageId;

                if (languageId > 0)
                {
                    yield return languageId;
                }
            }
        }

        public int GetCurrentLangId()
        {
            var result = CultureInfo.InstalledUICulture.LCID;

            var activeWindow = Application.ActiveWindow;
            var selection = activeWindow.Selection;
            var slides = selection.SlideRange;

            foreach (var slide in slides)
            {
                if (slide.Shapes != default)
                {
                    foreach (var shape in slide.Shapes)
                    {
                        if (shape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            result = (int)shape.TextFrame.TextRange.LanguageID;
                            break;
                        }
                    }
                }
            }

            return result;
        }

        public bool GetLanguageSetterExists(IRibbonControl control)
        {
            return this.TaskPaneExists(TaskPaneId);
        }

        public bool GetLanguageSetterPressed(IRibbonControl control)
        {
            return this.TaskPaneVisible(TaskPaneId);
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

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.RegisterInstance<ILanguageSetter>(this);

            var settingsPath = GetSettingsPath();
            var settings = new ConfigurationBuilder<ILanguageSettings>()
                .UseJsonFile(settingsPath).Build();

            containerRegistry.RegisterInstance(settings);
        }

        public void SetPresentationLanguage(int languageId)
        {
            var activePresentation = Application.ActivePresentation;

            activePresentation.DefaultLanguageID = (MsoLanguageID)languageId;

            var slides = activePresentation.Slides;

            foreach (var slide in slides)
            {
                SetSlideLanguage(
                    slide: slide,
                    languageId: languageId);
            }

            SetMasterLanguage(
                master: activePresentation.SlideMaster,
                languageId: languageId);

            if (activePresentation.HasTitleMaster == MsoTriState.msoTrue)
            {
                SetMasterLanguage(
                    master: activePresentation.TitleMaster,
                    languageId: languageId);
            }

            if (activePresentation.HasNotesMaster)
            {
                SetMasterLanguage(
                    master: activePresentation.NotesMaster,
                    languageId: languageId);
            }

            if (activePresentation.HasHandoutMaster)
            {
                SetMasterLanguage(
                    master: activePresentation.HandoutMaster,
                    languageId: languageId);
            }

            this.SetTaskPaneVisible(
                id: TaskPaneId,
                isVisible: false);
        }

        public void SetSlidesLanguage(int languageId)
        {
            var activeWindow = Application.ActiveWindow;
            var selection = activeWindow.Selection;
            var slides = selection.SlideRange;

            foreach (var slide in slides)
            {
                SetSlideLanguage(
                    slide: slide,
                    languageId: languageId);
            }

            this.SetTaskPaneVisible(
                id: TaskPaneId,
                isVisible: false);
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

        private static void SetShapeLanguage(NetOffice.PowerPointApi.Shape shape, int languageId)
        {
            if (shape.HasTextFrame == MsoTriState.msoTrue)
            {
                shape.TextFrame.TextRange.LanguageID = (MsoLanguageID)languageId;
            }

            if (shape.HasTable == MsoTriState.msoTrue)
            {
                var table = shape.Table;

                for (var row = 1; row < table.Rows.Count; row++)
                {
                    for (var column = 0; column < table.Columns.Count; column++)
                    {
                        var cell = table.Cell(
                            row: row,
                            column: column);

                        cell.Shape.TextFrame.TextRange.LanguageID = (MsoLanguageID)languageId;
                    }
                }
            }

            if (shape.Type == MsoShapeType.msoGroup || shape.Type == MsoShapeType.msoSmartArt)
            {
                foreach (var groupItem in shape.GroupItems)
                {
                    SetShapeLanguage(
                        shape: groupItem,
                        languageId: languageId);
                }
            }
        }

        private void OnAddInStart(ref Array custom)
        {
            Application.AfterNewPresentationEvent += OnPresentationNew;
            Application.AfterPresentationOpenEvent += OnPresentationOpen;
            Application.PresentationCloseEvent += OnPresentationClose;
        }

        private void OnPresentationClose(Presentation pres)
        {
            this.SetTaskPaneVisible(
                id: TaskPaneId,
                isVisible: false);
        }

        private void OnPresentationNew(Presentation pres)
        {
            OnLanguageUpdateEvent?.Invoke(
                sender: this,
                e: default);
        }

        private void OnPresentationOpen(Presentation pres)
        {
            OnLanguageUpdateEvent?.Invoke(
                sender: this,
                e: default);
        }

        private void SetMasterLanguage(object master, int languageId)
        {
            var relevant = master as Master;

            if (relevant != default)
            {
                foreach (var relevantShape in relevant.Shapes)
                {
                    SetShapeLanguage(
                        shape: relevantShape,
                        languageId: languageId);
                }

                try
                {
                    foreach (var customLayout in relevant.CustomLayouts)
                    {
                        var relevantLayout = customLayout as CustomLayout;

                        foreach (var relevantShape in relevant.Shapes)
                        {
                            SetShapeLanguage(
                                shape: relevantShape,
                                languageId: languageId);
                        }
                    }
                }
                catch (PropertyGetCOMException)
                { }
            }
        }

        private void SetSlideLanguage(object slide, int languageId)
        {
            var relevant = slide as Slide;

            if (relevant != default)
            {
                foreach (var relevantShape in relevant.Shapes)
                {
                    SetShapeLanguage(
                        shape: relevantShape,
                        languageId: languageId);
                }
            }
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Do not catch general exception types