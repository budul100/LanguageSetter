using System;
using System.Windows;
using System.Windows.Controls;

namespace LanguageView.Helpers
{
    public class ListBoxHelper
        : DependencyObject
    {
        #region Public Fields

        // Using a DependencyProperty as the backing store for AutoSizeItemCount.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty AutoSizeItemCountProperty = DependencyProperty
            .RegisterAttached("AutoSizeItemCount", typeof(int), typeof(ListBoxHelper), new PropertyMetadata(0, OnAutoSizeItemCountChanged));

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

        private static void OnAutoSizeItemCountChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var listBox = d as ListBox;

            listBox.AddHandler(
                routedEvent: ScrollViewer.ScrollChangedEvent,
                handler: new ScrollChangedEventHandler((lb, arg) => UpdateSize(listBox)));

            listBox.ItemContainerGenerator.ItemsChanged += (ig, arg) => UpdateSize(listBox);
        }

        private static void UpdateSize(ListBox listBox)
        {
            var gen = listBox.ItemContainerGenerator;

            if (listBox.InputHitTest(new Point(listBox.Padding.Left + 5, listBox.Padding.Top + 5)) is FrameworkElement element && gen != default)
            {
                var item = element.DataContext;

                if (item != default)
                {
                    if (!(gen.ContainerFromItem(item) is FrameworkElement container))
                    {
                        container = element;
                    }

                    var maxCount = GetAutoSizeItemCount(listBox);

                    var newHeight = Math.Min(maxCount, gen.Items.Count) * container.ActualHeight;
                    newHeight += listBox.Padding.Top + listBox.Padding.Bottom + listBox.BorderThickness.Top + listBox.BorderThickness.Bottom + 2;

                    if (listBox.ActualHeight != newHeight)
                    {
                        listBox.Height = newHeight;
                    }
                }
            }
        }

        #endregion Private Methods
    }
}