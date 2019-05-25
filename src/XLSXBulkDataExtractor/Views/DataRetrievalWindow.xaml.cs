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
        DataRetrievalViewModel _vm;
        public DataRetrievalWindow(DataRetrievalViewModel vm)
        {
            InitializeComponent();
            DataContext = _vm = vm;
        }
    }
}
