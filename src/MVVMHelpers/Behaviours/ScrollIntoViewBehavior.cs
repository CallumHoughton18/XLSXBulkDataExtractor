using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Interactivity;

namespace MVVMHelpers.Behaviours
{
    public class ScrollIntoViewBehavior : Behavior<ListView>
    {
        protected override void OnAttached()
        {
            ListView listBox = AssociatedObject;
            ((INotifyCollectionChanged)listBox.Items).CollectionChanged += OnListBox_CollectionChanged;
        }

        protected override void OnDetaching()
        {
            ListView listBox = AssociatedObject;
            ((INotifyCollectionChanged)listBox.Items).CollectionChanged -= OnListBox_CollectionChanged;
        }

        private void OnListBox_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            ListView listBox = AssociatedObject;
            if (e.Action == NotifyCollectionChangedAction.Add) listBox.ScrollIntoView(e.NewItems[0]);         
        }
    }
}
