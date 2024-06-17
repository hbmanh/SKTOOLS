using System.Collections;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace SKToolsAddins.Utils
{
    public class MultiSelect
    {
        #region BindableSelectedItemsProperty
        public static readonly DependencyProperty BindableSelectedItemsProperty
            = DependencyProperty.RegisterAttached("BindableSelectedItems", typeof(IList),
                typeof(MultiSelect),
                new PropertyMetadata(default, OnBindableSelectedItemsChanged));

        public static IList GetBindableSelectedItems(DependencyObject element) => (IList)element.GetValue(BindableSelectedItemsProperty);
        public static void SetBindableSelectedItems(DependencyObject element, IList value) => element.SetValue(BindableSelectedItemsProperty, value);
        #endregion

        #region OnBindableSelectedItemsChanged
        private static void OnBindableSelectedItemsChanged(DependencyObject element, DependencyPropertyChangedEventArgs args)
        {
            if (!(args.NewValue is IList newItems))
                return;

            if ((element is MultiSelector multiSelector))
            {
                multiSelector.SelectionChanged -= SelectorOnSelectionChanged;
                multiSelector.SelectionChanged += SelectorOnSelectionChanged;

                if (multiSelector is DataGrid grid)
                {
                    if (grid.SelectionMode == DataGridSelectionMode.Single)
                    {
                        multiSelector.SelectedItem = newItems.Count > 0 ? newItems[0] : null;
                    }
                    else
                    {
                        multiSelector.SelectedItems.Clear();

                        foreach (var newItem in newItems)
                        {
                            multiSelector.SelectedItems.Add(newItem);
                        }
                    }
                }
            }
            else if (element is Selector selector)
            {
                selector.SelectionChanged -= SelectorOnSelectionChanged;
                selector.SelectionChanged += SelectorOnSelectionChanged;

                // Selector is supported ListBox only:
                if (selector is ListBox listBox)
                {
                    if (listBox.SelectionMode == SelectionMode.Single)
                    {
                        selector.SelectedItem = newItems.Count > 0 ? newItems[0] : null;
                    }
                    else
                    {
                        listBox.SelectedItems.Clear();
                        foreach (var newItem in newItems)
                        {
                            listBox.SelectedItems.Add(newItem);
                        }
                    }
                }
            }
            else
            {
                return;
            }
        }
        #endregion

        #region SelectorOnSelectionChanged
        private static void SelectorOnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            IList viewModelSelectedItemList;

            if (sender is MultiSelector multiSelector)
            {
                viewModelSelectedItemList = GetBindableSelectedItems(multiSelector);
            }
            else if (sender is Selector selector)
            {
                viewModelSelectedItemList = GetBindableSelectedItems(selector);
            }
            else
            {
                viewModelSelectedItemList = null;
            }

            if (viewModelSelectedItemList == null)
                return;

            foreach (var item in e.AddedItems)
            {
                if (viewModelSelectedItemList.Contains(item))
                    continue;
                viewModelSelectedItemList.Add(item);
            }

            foreach (var eRemovedItem in e.RemovedItems)
            {
                if (!viewModelSelectedItemList.Contains(eRemovedItem))
                    continue;
                viewModelSelectedItemList.Remove(eRemovedItem);

            }
        }
        #endregion
    }
}
