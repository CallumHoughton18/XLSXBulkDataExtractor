using System.Windows;
using System.Windows.Controls;
using XLSXBulkDataExtractor.WPFLogic.Interfaces;
using XLSXBulkDataExtractor.WPFLogic.ViewModels;

namespace XLSXBulkDataExtractor
{
    /// <summary>
    /// Interaction logic for DataRetrievalWindow.xaml
    /// </summary>
    public partial class DataRetrievalWindow : Window
    {
        DataRetrievalViewModel vm;
        public DataRetrievalWindow(IIOService ioService, IXLIOService xlioService, IUIControlsService uiControlsService)
        {
            InitializeComponent();
            DataContext = vm = new DataRetrievalViewModel(ioService, xlioService, uiControlsService);
        }

        private void ListView_LayoutUpdated(object sender, System.EventArgs e)
        {
            ListView listView = sender as ListView;

            if (listView != null)
            {
                listView.ScrollIntoView(listView.Items[listView.Items.Count-1]);
            }
        }
    }
}
