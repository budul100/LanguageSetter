using PrismTaskPanes.Extensions;
using System;
using System.Windows;
using System.Windows.Controls;

namespace LanguageModule.Views
{
    public partial class LanguageView
        : UserControl
    {
        #region Public Constructors

        public LanguageView()
        {
            // This pre-load is necessary for the trigger in LanguageView only
            // See https://github.com/microsoft/XamlBehaviorsWpf/issues/86

            var _ = new Microsoft.Xaml.Behaviors.DefaultTriggerAttribute(
                targetType: typeof(Trigger),
                triggerType: typeof(Microsoft.Xaml.Behaviors.TriggerBase),
                parameters: null);

            try
            {
                InitializeComponent();
            }
            catch (Exception)
            {
                this.LoadViewFromUri("/LanguageModule;component/views/languageview.xaml");
            }
        }

        #endregion Public Constructors
    }
}