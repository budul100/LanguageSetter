using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace LanguageModule.Helpers
{
    public class ListBoxHelper
        : DependencyObject
    {
        #region Public Fields

        // Using a DependencyProperty as the backing store for AutoSizeItemCount.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty AutoSizeItemCountProperty = DependencyProperty.RegisterAttached(
            name: "AutoSizeItemCount",
            propertyType: typeof(int),
            ownerType: typeof(ListBoxHelper),
            defaultMetadata: new PropertyMetadata(0, OnAutoSizeItemCountChanged));

        #endregion Public Fields

        #region Public Methods

        public static int GetAutoSizeItemCount(DependencyObject obj)
        {
            return (int)obj.GetValue(AutoSizeItemCountProperty);
        }

        public static void SetAutoSizeItemCount(DependencyObject obj, int value)
        {
            obj.SetValue(
                dp: AutoSizeItemCountProperty,
                value: value);
        }

        #endregion Public Methods

        #region Private Methods

        private static T GetChildOfType<T>(DependencyObject depObj)
            where T : DependencyObject
        {
            var result = default(T);

            if (depObj != default)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    var child = VisualTreeHelper.GetChild(depObj, i);
                    result = (child as T) ?? GetChildOfType<T>(child);

                    if (result != default) break;
                }
            }

            return result;
        }

        private static void OnAutoSizeItemCountChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var listBox = d as ListBox;

            // we set this to 0.0 so that we ddon't create any elements
            // before we have had a chance to modify the scrollviewer
            listBox.MaxHeight = 0.0;

            listBox.AddHandler(
                routedEvent: ScrollViewer.ScrollChangedEvent,
                handler: new ScrollChangedEventHandler((lb, arg) => UpdateSize(listBox)));

            listBox.ItemContainerGenerator.ItemsChanged += (ig, arg) => UpdateSize(listBox);
        }

        private static void OnVirtualizingStackPanelSizeChanged(object sender, SizeChangedEventArgs e)
        {
            var stackPanel = sender as VirtualizingStackPanel;

            var listBox = (ListBox)ItemsControl.GetItemsOwner(stackPanel);
            var maxCount = GetAutoSizeItemCount(listBox);

            stackPanel.ScrollOwner.MaxHeight = stackPanel.Children.Count == 0
                ? 1
                : ((FrameworkElement)stackPanel.Children[0]).ActualHeight * maxCount;
        }

        private static void UpdateSize(ListBox listBox)
        {
            var scrollViewer = GetChildOfType<ScrollViewer>(listBox);

            if (scrollViewer != default)
            {
                // limit the scrollviewer height so that the bare minimum elements are generated
                scrollViewer.MaxHeight = 1.0;

                var stackPanel = GetChildOfType<VirtualizingStackPanel>(listBox);
                if (stackPanel != default)
                {
                    stackPanel.SizeChanged += OnVirtualizingStackPanelSizeChanged;
                }
            }

            listBox.MaxHeight = double.PositiveInfinity;
        }

        #endregion Private Methods
    }
}