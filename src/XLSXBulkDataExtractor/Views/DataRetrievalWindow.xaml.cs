using System.Windows;
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
    }
}
